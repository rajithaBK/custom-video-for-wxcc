import "./style.css";
import getAccessToken from "./utils/get-access-token.js";
import getGuestToken from "./utils/get-guest-token.js";
// import getDestLinks from "./utils/get-dest-links.js";
import getDestLinks from "./utils/get-dest-links-whiteboarding.js";
require("dotenv").config();
const {
  enableDrag
} = require("./selfview");
var socket = io();
const axios = require("axios");
console.log("socketio", socket);

// Change this value to the Live Chat App used in the WxCC Connect tenant
const connectChatLiveChatAppId = 'VI24093513';
const urlParams = new URLSearchParams(window.location.search);
const name = urlParams.get("name") !== null ? urlParams.get("name") : "John Doe";
const email = urlParams.get("email") !== null ? urlParams.get("email") : "johndoe@gmail.com";
const SAAccessToken = urlParams.get("access_token") !== null ? urlParams.get("access_token") : await getAccessToken();
console.log("got access token", SAAccessToken);
const myAccessToken = SAAccessToken;
console.log("got guest token", myAccessToken);
const destination = await getDestLinks(SAAccessToken);
console.log("got dest links");

// const wc_body = {
const wc_body = JSON.stringify({
  customerName: name,
  customerEmail: email,
  videoCallDestination: destination,
  "inappmessaging.appId": connectChatLiveChatAppId,
  "inappmessaging.userId": "6806ea7s-a04e-4fdb-9d86-0b33626f3577"
});
let config = {
  method: 'post',
  url: 'https://hooks.us.webexconnect.io/events/DJ03LZTWMI',
  headers: {
    'Content-Type': 'application/json'
  },
  data: wc_body
};
await axios.request(config)
// await axios
// .post("https://hooks.us.webexconnect.io/events/F8JF7AK1LS", wc_body)
.then(resp => {
  console.log("WC Success");
});
if (!myAccessToken) {
  alert("Access token is missing. ");
  throw new Error("Access token is missing. ");
}
const Webex = require("webex");
const webex = Webex.init({
  credentials: {
    access_token: myAccessToken
  }
});
webex.config.logger.level = "debug";
webex.meetings.register().then(r => {
  console.log("Succesfully registered");
  console.log(destination);
  webex.meetings.create(destination).then(async meeting => {
    console.log("Meeting successfully created");

    // Call our helper function for binding events to meetings
    await bindMeetingEvents(meeting);
    await bindButtonEvents(meeting);
    await joinMeeting(meeting);
  }).catch(error => {
    // Report the error
    console.error(error);
  });
}).catch(err => {
  console.error(err);
  alert(err);
  throw err;
});
async function bindButtonEvents(meeting) {
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
  const meetingDest = meeting.destination;
  document.getElementById("hangup").addEventListener("click", async () => {
    // window.location.href = "/hangup";
    console.log("hangup clicked");
    await meeting.leave();
    await meeting.getMembers().then(members => {
      console.log("members", members);
    });
  });
  videoMuteOff.addEventListener("click", () => {
    console.log("videmute off clicked");
    meeting.muteVideo();
    videoMuteOff.style.display = "none";
    videoMuteOn.style.display = "";
  });
  videoMuteOn.addEventListener("click", () => {
    console.log("video mute on clicked");
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
    // dropdown.classList.toggle("is-active");
    Array.from(dropdown).forEach(drop => {
      drop.classList.toggle("is-active");
    });
  });
  if (self) {
    enableDrag(self, remoteView);
  }

  // Handle orientation changes
  window.addEventListener('resize', handleOrientationChange);
  if (screen.orientation) {
    screen.orientation.addEventListener('change', handleOrientationChange);
  }
  // Set initial orientation
  handleOrientationChange();
}
async function bindMeetingEvents(meeting) {
  const selfView = document.getElementById("self-view");
  const remoteViewVideo = document.getElementById("remote-view-video");
  const remoteViewAudio = document.getElementById("remote-view-audio");
  const buttonsContainer = document.getElementById("buttons-container");
  const loadingContainer = document.getElementById("loading-container");
  meeting.on("error", error => console.log(error, "Meeting Error"));
  meeting.on("media:ready", media => {
    if (!media) return;
    const element = media.type === "local" ? selfView : media.type === "remoteVideo" ? remoteViewVideo : media.type === "remoteAudio" ? remoteViewAudio : null;
    if (element) {
      element.srcObject = media.stream;
      buttonsContainer.style.display = "flex";
      loadingContainer.style.display = "none";
    }
  });
  meeting.on("media:stopped", media => {
    console.log("meeting stopped");
    meeting.stopRecording();
    webex.meetings.unregister();
    window.location.href = "hangup.html";
    const element = media.type === "local" ? selfView : media.type === "remoteVideo" ? remoteViewVideo : media.type === "remoteAudio" ? remoteViewAudio : null;
    if (element) {
      element.srcObject = null;
      buttonsContainer.style.display = "none";
    }
  });
}

// Join the meeting and add media
// Join the meeting and add media through joinWithMedia method.
async function joinMeeting(meeting) {
  try {
    const {
      sendAudio,
      sendVideo
    } = await meeting.getSupportedDevices({
      sendAudio: true,
      sendVideo: true
    });
    console.log('SendAudio, sendvideo;', sendAudio, sendVideo);
    meeting.join().then(async () => {
      const mediaSettings = {
        receiveVideo: true,
        receiveAudio: true,
        receiveShare: false,
        sendShare: false,
        sendVideo,
        sendAudio
      };

      // Get our local media stream and add it to the meeting
      meeting.getMediaStreams(mediaSettings).then(mediaStreams => {
        const [localStream, localShare] = mediaStreams;
        meeting.addMedia({
          localShare,
          localStream,
          mediaSettings
        });
      });
    }).catch(error => {
      console.log('meeting.join error:', error);
    });
  } catch (error) {
    console.log(error, "Join Meeting Error");
    throw error;
  }
}
function handleOrientationChange() {
  const self = document.getElementById('self');
  const isLandscape = window.innerWidth > window.innerHeight;
  if (isLandscape) {
    self.style.width = '40%';
    self.style.height = 'auto';
  } else {
    self.style.width = '25%';
    self.style.height = 'auto';
  }
}