// WxCC "video-in-queue" widget entry (browser-native).
//
// Flow when a request lands on the agent desktop:
//   1) POST the agent's WxCC access token to our backend (/api/session). The
//      backend creates a shared Webex meeting, mints a guest token for the
//      agent, and auto-dials the fixed kiosk (Desk Pro G2) into that meeting.
//   2) This page joins the SAME meeting as a guest using the returned token.
//
// No secrets live here. `Webex` and `axios` are provided as globals by the
// CDN <script> tags in index.html (webex UMD + axios).

const BACKEND_URL = "https://wxcc-video-backend.vercel.app/api/session";

const urlParams = new URLSearchParams(window.location.search);
// The wrapper (video-cc-widget.js) injects the agent token as ?access_token=...
const agentToken = urlParams.get("access_token");

const loadingContainer = document.getElementById("loading-container");

function setStatus(msg) {
  if (loadingContainer) loadingContainer.innerHTML = `<span>${msg}</span>`;
}
// Pull a human-readable reason out of whatever the SDK / axios throws so the
// UI shows the real cause instead of a generic message.
function errorDetail(err) {
  if (!err) return "";
  const parts = [];
  if (err.response && err.response.data) {
    const d = err.response.data;
    parts.push(d.error || d.detail || d.message || JSON.stringify(d));
  }
  if (err.body) {
    parts.push(typeof err.body === "string" ? err.body : JSON.stringify(err.body));
  }
  if (err.message) parts.push(err.message);
  if (!parts.length && typeof err === "string") parts.push(err);
  if (!parts.length) {
    try {
      parts.push(JSON.stringify(err));
    } catch (_) {
      parts.push(String(err));
    }
  }
  return parts.filter(Boolean).join(" — ");
}

function fail(msg, err) {
  const detail = errorDetail(err);
  const full = detail ? `${msg} (${detail})` : msg;
  if (err) console.error(msg, err);
  else console.error(msg);
  setStatus(full);
}

let webex;
let currentMeeting = null;

async function start() {
  if (!agentToken || String(agentToken).indexOf("$STORE") !== -1) {
    fail("Access token is missing.");
    return;
  }

  setStatus("Connecting…");

  // 1) Ask the backend to set up the session (meeting + guest token + kiosk dial).
  let session;
  try {
    const resp = await axios.post(
      BACKEND_URL,
      {},
      { headers: { Authorization: `Bearer ${agentToken}` } }
    );
    session = resp.data;
  } catch (e) {
    fail("Failed to start the video session.", e);
    return;
  }

  // The backend returns the SA#1 (meeting owner) access token. Joining with it
  // makes the agent the HOST, which starts the meeting so the auto-dialed kiosk
  // is admitted immediately — no lobby, no "waiting for host".
  const { agentJoinToken, meetingSip, meetingPassword, dialStatus } =
    session || {};
  console.log("session:", { meetingSip, dialStatus });
  if (dialStatus && dialStatus !== "ok") {
    // Non-fatal: the agent can still join; the kiosk may need to join manually.
    console.warn("Kiosk auto-dial did not succeed:", dialStatus);
  }
  if (!agentJoinToken || !meetingSip) {
    fail("Invalid session response from backend.");
    return;
  }

  // 2) Join the SAME meeting as the host (using the owner's access token).
  //
  //   a) WAIT for the SDK `ready` event before calling meetings.register().
  //      Calling register() early is what caused the old
  //      "Cannot read properties of undefined (reading 'internal')" crash.
  //   b) As the host/owner the SDK does NOT require the meeting password, so
  //      verifyPassword() is only a guarded fallback: we attempt it if a
  //      password is present but never hard-fail the join if it errors.
  let meeting;
  try {
    webex = window.Webex.init({ credentials: { access_token: agentJoinToken } });
    if (webex.config && webex.config.logger) webex.config.logger.level = "debug";

    await waitForReady(webex);
  } catch (err) {
    fail("Could not initialize the video client.", err);
    return;
  }

  try {
    await webex.meetings.register();
    console.log("Webex meetings registered");
  } catch (err) {
    // Most likely cause: the join token lacks the meetings scopes.
    fail("Could not register the video client.", err);
    return;
  }

  try {
    meeting = await webex.meetings.create(meetingSip);
    currentMeeting = meeting;
    console.log("Meeting created for", meetingSip);
  } catch (err) {
    fail("Could not look up the meeting.", err);
    return;
  }

  // Guarded fallback only — the host token doesn't need the password, and the
  // SDK throws "password was not required" when we pass one as the owner.
  if (meetingPassword) {
    try {
      const vp = await meeting.verifyPassword(meetingPassword);
      console.log("verifyPassword:", vp);
    } catch (e) {
      console.warn("verifyPassword skipped/failed (continuing):", e);
    }
  }

  try {
    await bindMeetingEvents(meeting);
    bindButtonEvents(meeting);
    bindLobbyEvents(meeting);
    await joinMeeting(meeting);
  } catch (err) {
    fail("Could not join the meeting.", err);
  }
}

// If the SDK ever reports the self participant is waiting in the lobby (e.g. if
// the join identity is changed back to a non-host guest), show a clear status
// instead of leaving the generic "Connecting…" or a misleading error.
function bindLobbyEvents(meeting) {
  try {
    meeting.on("meeting:self:lobbyWaiting", () => {
      console.log("self is waiting in the lobby");
      setStatus("Waiting to be admitted…");
    });
    meeting.on("meeting:self:guestAdmitted", () => {
      console.log("self was admitted from the lobby");
      setStatus("Admitted — connecting media…");
    });
  } catch (e) {
    console.warn("could not bind lobby events", e);
  }
}

// Resolves once the Webex SDK has finished initializing. Registering meetings
// before this fires throws deep inside the SDK ("reading 'internal'").
function waitForReady(webexInstance) {
  return new Promise((resolve, reject) => {
    if (webexInstance.ready) return resolve();
    const timeout = setTimeout(
      () => reject(new Error("Webex SDK ready timeout")),
      20000
    );
    webexInstance.once("ready", () => {
      clearTimeout(timeout);
      resolve();
    });
  });
}

function bindButtonEvents(meeting) {
  const videoMuteOff = document.getElementById("video-mute-off");
  const videoMuteOn = document.getElementById("video-mute-on");
  const audioMuteOff = document.getElementById("audio-mute-off");
  const audioMuteOn = document.getElementById("audio-mute-on");
  const dropdownButton = document.getElementById("dropdown-button");
  const hideSelfView = document.getElementById("hide-self-view");
  const showSelfView = document.getElementById("show-self-view");
  const self = document.getElementById("self");
  const remoteView = document.getElementById("remote-view");
  const dropdown = document.getElementsByClassName("dropdown");

  document.getElementById("hangup").addEventListener("click", async () => {
    console.log("hangup clicked");
    try {
      await meeting.leave();
    } catch (e) {
      console.error("leave error", e);
    }
  });

  videoMuteOff.addEventListener("click", () => {
    meeting.muteVideo();
    videoMuteOff.style.display = "none";
    videoMuteOn.style.display = "";
  });
  videoMuteOn.addEventListener("click", () => {
    meeting.unmuteVideo();
    videoMuteOff.style.display = "";
    videoMuteOn.style.display = "none";
  });
  audioMuteOff.addEventListener("click", () => {
    meeting.muteAudio();
    audioMuteOff.style.display = "none";
    audioMuteOn.style.display = "";
  });
  audioMuteOn.addEventListener("click", () => {
    meeting.unmuteAudio();
    audioMuteOff.style.display = "";
    audioMuteOn.style.display = "none";
  });

  hideSelfView.addEventListener("click", () => {
    self.style.display = "none";
    hideSelfView.style.display = "none";
    showSelfView.style.display = "";
  });
  showSelfView.addEventListener("click", () => {
    self.style.display = "";
    hideSelfView.style.display = "";
    showSelfView.style.display = "none";
  });

  dropdownButton.addEventListener("click", () => {
    Array.from(dropdown).forEach((drop) => drop.classList.toggle("is-active"));
  });

  if (self) enableDrag(self, remoteView);

  window.addEventListener("resize", handleOrientationChange);
  if (screen.orientation) {
    screen.orientation.addEventListener("change", handleOrientationChange);
  }
  handleOrientationChange();
}

async function bindMeetingEvents(meeting) {
  const selfView = document.getElementById("self-view");
  const remoteViewVideo = document.getElementById("remote-view-video");
  const remoteViewAudio = document.getElementById("remote-view-audio");
  const buttonsContainer = document.getElementById("buttons-container");

  meeting.on("error", (error) => console.log(error, "Meeting Error"));

  // In webex 3.x, 'media:ready' delivers the REMOTE streams (remoteVideo /
  // remoteAudio / remoteShare). The LOCAL self-view is attached separately in
  // joinMeeting() from the created camera stream's `.outputStream`.
  meeting.on("media:ready", (media) => {
    if (!media || !media.stream) return;
    const element =
      media.type === "remoteVideo"
        ? remoteViewVideo
        : media.type === "remoteAudio"
        ? remoteViewAudio
        : media.type === "local"
        ? selfView
        : null;
    if (element) {
      element.srcObject = media.stream;
      if (buttonsContainer) buttonsContainer.style.display = "flex";
      if (loadingContainer) loadingContainer.style.display = "none";
    }
  });

  meeting.on("media:stopped", (media) => {
    if (!media) return;
    const element =
      media.type === "remoteVideo"
        ? remoteViewVideo
        : media.type === "remoteAudio"
        ? remoteViewAudio
        : media.type === "local"
        ? selfView
        : null;
    if (element) element.srcObject = null;
  });

  // Full teardown when the local user leaves / the meeting ends.
  meeting.on("meeting:self:left", cleanupAndRedirect);
  meeting.on("meeting:ended", cleanupAndRedirect);
}

let __tornDown = false;
function cleanupAndRedirect() {
  if (__tornDown) return;
  __tornDown = true;
  console.log("meeting ended — cleaning up");
  try {
    const streams = currentMeeting && currentMeeting.__localStreams;
    if (streams) {
      if (streams.microphone && streams.microphone.stop) streams.microphone.stop();
      if (streams.camera && streams.camera.stop) streams.camera.stop();
    }
  } catch (e) {}
  try {
    webex.meetings.unregister();
  } catch (e) {}
  window.location.href = "hangup.html";
}

// Webex JS SDK v3 (3.7.0) media flow.
//
// The old v2 API (getSupportedDevices / getMediaStreams /
// addMedia({localStream, localShare})) does NOT exist in 3.x and throws
// "meeting.getSupportedDevices is not a function". In 3.x you:
//   1) create local streams via webex.meetings.mediaHelpers
//      (createMicrophoneStream / createCameraStream),
//   2) meeting.join({ enableMultistream: false })  (single-stream keeps the
//      simple media:ready path for remote audio/video),
//   3) meeting.addMedia({ localStreams: { microphone, camera },
//                         audioEnabled, videoEnabled }).
// Local video is shown from the created camera stream's `.outputStream`
// (a MediaStream); remote media arrives via the 'media:ready' event.
async function joinMeeting(meeting) {
  const mediaHelpers = webex.meetings.mediaHelpers;

  // Create local streams up front, but never hard-fail if a device is missing
  // or permission is denied — audio-only, or even receive-only, should still
  // connect instead of throwing.
  let microphone = null;
  let camera = null;
  try {
    microphone = await mediaHelpers.createMicrophoneStream();
  } catch (e) {
    console.warn("Microphone unavailable (continuing without mic):", e && e.name, e && e.message);
  }
  try {
    camera = await mediaHelpers.createCameraStream();
  } catch (e) {
    console.warn("Camera unavailable (continuing without camera):", e && e.name, e && e.message);
  }

  // Keep references for cleanup on hangup.
  meeting.__localStreams = { microphone, camera };

  // Show the local camera immediately in the self-view.
  if (camera && camera.outputStream) {
    const selfView = document.getElementById("self-view");
    if (selfView) selfView.srcObject = camera.outputStream;
    const buttonsContainer = document.getElementById("buttons-container");
    if (buttonsContainer) buttonsContainer.style.display = "flex";
    if (loadingContainer) loadingContainer.style.display = "none";
  }

  // 1) Join (single-stream so remote media surfaces via media:ready).
  try {
    await meeting.join({ enableMultistream: false });
    console.log("meeting joined");
  } catch (error) {
    fail("Could not join the meeting.", error);
    throw error;
  }

  // 2) Publish local media / negotiate remote media.
  const localStreams = {};
  if (microphone) localStreams.microphone = microphone;
  if (camera) localStreams.camera = camera;
  try {
    await meeting.addMedia({
      localStreams,
      audioEnabled: !!microphone,
      videoEnabled: !!camera,
    });
    console.log("media added");
  } catch (error) {
    // We're in the meeting; only media negotiation failed. Surface the real
    // reason but don't treat it as a failure to join.
    fail("Joined, but could not set up audio/video.", error);
  }
}

function handleOrientationChange() {
  // Self-view sizing/position is handled by style.css now. This is kept so the
  // resize/orientation listeners bound in bindButtonEvents remain valid no-ops.
}

// Minimal draggable self-view (replaces the bundled ./selfview helper).
function enableDrag(el, bounds) {
  let dragging = false;
  let startX = 0;
  let startY = 0;
  let origLeft = 0;
  let origTop = 0;

  const onDown = (e) => {
    dragging = true;
    const p = e.touches ? e.touches[0] : e;
    startX = p.clientX;
    startY = p.clientY;
    const rect = el.getBoundingClientRect();
    const parent = (bounds || el.parentElement).getBoundingClientRect();
    origLeft = rect.left - parent.left;
    origTop = rect.top - parent.top;
    el.style.position = "absolute";
    e.preventDefault();
  };
  const onMove = (e) => {
    if (!dragging) return;
    const p = e.touches ? e.touches[0] : e;
    el.style.left = origLeft + (p.clientX - startX) + "px";
    el.style.top = origTop + (p.clientY - startY) + "px";
  };
  const onUp = () => {
    dragging = false;
  };

  el.addEventListener("mousedown", onDown);
  el.addEventListener("touchstart", onDown, { passive: false });
  window.addEventListener("mousemove", onMove);
  window.addEventListener("touchmove", onMove, { passive: false });
  window.addEventListener("mouseup", onUp);
  window.addEventListener("touchend", onUp);
}

start();
