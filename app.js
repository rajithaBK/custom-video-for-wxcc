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
function fail(msg, err) {
  if (err) console.error(msg, err);
  else console.error(msg);
  setStatus(msg);
}

let webex;

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

  const { agentGuestToken, meetingSip, meetingPassword, dialStatus } =
    session || {};
  console.log("session:", { meetingSip, dialStatus });
  if (dialStatus && dialStatus !== "ok") {
    // Non-fatal: the agent can still join; the kiosk may need to join manually.
    console.warn("Kiosk auto-dial did not succeed:", dialStatus);
  }
  if (!agentGuestToken || !meetingSip) {
    fail("Invalid session response from backend.");
    return;
  }

  // 2) Join the SAME meeting as a guest.
  //
  // The Service App guest token (POST /v1/guests/token) is meetings-capable.
  // Two things are essential for the guest join to work:
  //   a) WAIT for the SDK `ready` event before calling meetings.register().
  //      Calling register() early is what caused the old
  //      "Cannot read properties of undefined (reading 'internal')" crash.
  //   b) The site assigns a meeting password even for allowJoin meetings, so
  //      verify it (meeting.verifyPassword) before join() to avoid the
  //      "Password is required" rejection.
  try {
    webex = window.Webex.init({ credentials: { access_token: agentGuestToken } });
    if (webex.config && webex.config.logger) webex.config.logger.level = "debug";

    await waitForReady(webex);
    await webex.meetings.register();
    console.log("Webex meetings registered");

    const meeting = await webex.meetings.create(meetingSip);
    console.log("Meeting created for", meetingSip);

    if (meetingPassword) {
      try {
        const vp = await meeting.verifyPassword(meetingPassword);
        console.log("verifyPassword:", vp);
      } catch (e) {
        console.warn("verifyPassword failed (continuing):", e);
      }
    }

    await bindMeetingEvents(meeting);
    bindButtonEvents(meeting);
    await joinMeeting(meeting);
  } catch (err) {
    fail("Could not join the meeting.", err);
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

  meeting.on("media:ready", (media) => {
    if (!media) return;
    const element =
      media.type === "local"
        ? selfView
        : media.type === "remoteVideo"
        ? remoteViewVideo
        : media.type === "remoteAudio"
        ? remoteViewAudio
        : null;
    if (element) {
      element.srcObject = media.stream;
      buttonsContainer.style.display = "flex";
      if (loadingContainer) loadingContainer.style.display = "none";
    }
  });

  meeting.on("media:stopped", (media) => {
    console.log("meeting stopped");
    try {
      meeting.stopRecording();
    } catch (e) {}
    try {
      webex.meetings.unregister();
    } catch (e) {}
    window.location.href = "hangup.html";
    const element =
      media.type === "local"
        ? selfView
        : media.type === "remoteVideo"
        ? remoteViewVideo
        : media.type === "remoteAudio"
        ? remoteViewAudio
        : null;
    if (element) {
      element.srcObject = null;
      buttonsContainer.style.display = "none";
    }
  });
}

async function joinMeeting(meeting) {
  try {
    const { sendAudio, sendVideo } = await meeting.getSupportedDevices({
      sendAudio: true,
      sendVideo: true,
    });
    meeting
      .join()
      .then(async () => {
        const mediaSettings = {
          receiveVideo: true,
          receiveAudio: true,
          receiveShare: false,
          sendShare: false,
          sendVideo,
          sendAudio,
        };
        meeting.getMediaStreams(mediaSettings).then((mediaStreams) => {
          const [localStream, localShare] = mediaStreams;
          meeting.addMedia({ localShare, localStream, mediaSettings });
        });
      })
      .catch((error) => console.log("meeting.join error:", error));
  } catch (error) {
    fail("Join Meeting Error", error);
    throw error;
  }
}

function handleOrientationChange() {
  const self = document.getElementById("self");
  if (!self) return;
  const isLandscape = window.innerWidth > window.innerHeight;
  self.style.width = isLandscape ? "40%" : "25%";
  self.style.height = "auto";
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
