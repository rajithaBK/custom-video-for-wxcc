// Thin WxCC wrapper web component.
// WxCC resolves $STORE.* values for web-component *properties* (not for a
// raw iframe src), so we receive the agent access token here and pass it
// into the hosted video page via its ?access_token=... query parameter.
(function () {
  var PAGE_URL = "https://rajithabk.github.io/custom-video-for-wxcc/index.html";

  class VideoCcWidget extends HTMLElement {
    set accessToken(value) {
      this._accessToken = value;
      this._maybeRender();
    }
    get accessToken() {
      return this._accessToken;
    }

    connectedCallback() {
      this._connected = true;
      this.style.display = "block";
      this.style.height = "100%";
      this._maybeRender();
    }

    _maybeRender() {
      if (!this._connected || this._rendered) return;
      var token = this._accessToken;
      // Wait until WxCC has injected a real, non-placeholder token.
      if (!token || String(token).indexOf("$STORE") !== -1) return;

      var src = PAGE_URL + "?access_token=" + encodeURIComponent(token);
      var iframe = document.createElement("iframe");
      iframe.setAttribute("src", src);
      iframe.setAttribute("style", "width:100%;height:100%;border:none;");
      iframe.setAttribute(
        "allow",
        "camera; microphone; autoplay; display-capture"
      );
      this.appendChild(iframe);
      this._rendered = true;
    }
  }

  if (!customElements.get("video-cc-widget")) {
    customElements.define("video-cc-widget", VideoCcWidget);
  }
})();
