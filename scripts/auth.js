// 인증코드 확인. 레포디와 파워링크 노출 점검이 같은 파일을 복사해서 쓴다(고칠 때는 두 곳을 함께 고칠 것).
// 인터넷에 올린 해시(AUTH_URL)와 이 PC에 저장한 해시가 같으면 바로 연다. 다르면 인증코드를 묻는다.
// 인증코드 원문은 어디에도 저장하지 않고 해시만 쓴다. 이 PC의 기록은 두 프로그램이 함께 쓴다(한 번 입력하면 둘 다 열린다).
// 소스를 고치면 풀리는 잠금장치 수준이다. 무단 재배포, 퇴사 후 사용을 막는 용도.
// 인증은 한 번이지만 사용 기록(USAGE_FORM_URL)은 프로그램별로 보낸다. 새 프로그램도 authGate("이름")만 부르면 그 이름으로 쌓인다.
var AUTH_URL = "https://gist.githubusercontent.com/hojong90-creator/01d3fc76a5643d30841218cb5107cad7/raw/06be7ff16b0f1c16a5b4ded67005daafbbca78dd/auth.txt";               // Gist Raw 주소. 비어 있으면 검사하지 않는다
var AUTH_SALT = "ad-ops-toolkit|";
var AUTH_GRACE_DAYS = 7;         // 인터넷이 안 될 때는 마지막으로 확인한 뒤 이 기간까지 연다
var AUTH_TIMEOUT_MS = 10000;

// 사용 기록. 구글 설문지(도구 사용 기록, hjpak@ye-ji.com)에 응답을 보내 시트에 쌓는다. 비어 있으면 보내지 않는다
var USAGE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSfHGM7FdPs-6IGVyPjW5NSY-AUA-NgJBY_rFUDxyNfixlsLEQ/formResponse";
var USAGE_ENTRY = { program: "entry.1525903788", kind: "entry.1471018899", pc: "entry.55782683", user: "entry.985032950" };

var authLocked = false, authAppName = "";

function authTrim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }

// 인증코드 → 해시(64자리). 인증코드 만들기.hta도 이 함수를 쓴다.
function authHash(code) { return sha256Hex(AUTH_SALT + authTrim(code)); }

// Gist에서 복사한 주소를 항상 최신 내용을 읽는 주소로 바꾼다.
// Raw 주소의 중간 버전 번호(40자리)를 빼고, Gist 화면 주소면 뒤에 /raw를 붙인다.
function authRawUrl(u) {
  u = authTrim(u).replace(/\/raw\/[0-9a-f]{40}\//i, "/raw/");
  if (/^https:\/\/gist\.github\.com\/[^\/]+\/[0-9a-f]+\/?$/i.test(u)) { u = u.replace(/\/?$/, "/raw"); }
  return u;
}

function authStorePath() {
  var sh = new ActiveXObject("WScript.Shell");
  return sh.ExpandEnvironmentStrings("%APPDATA%") + "\\ad-ops-toolkit\\auth.json";
}
function authLoad() {
  try {
    var f = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(authStorePath(), 1);
    var o = JSON.parse(f.ReadAll()); f.Close();
    return { Hash: String(o.Hash || ""), CheckedAt: Number(o.CheckedAt) || 0 };
  } catch (e) { return { Hash: "", CheckedAt: 0 }; }
}
function authSave(hash) {
  try {
    var fso = new ActiveXObject("Scripting.FileSystemObject"), p = authStorePath(), dir = fso.GetParentFolderName(p);
    if (!fso.FolderExists(dir)) { fso.CreateFolder(dir); }
    var f = fso.CreateTextFile(p, true);
    f.Write(JSON.stringify({ Hash: hash, CheckedAt: new Date().getTime() })); f.Close();
  } catch (e) {}
}

// 인터넷의 해시를 읽는다. done(해시) 또는 done(null, 이유)
function authFetch(done) {
  var x = null, finished = false;
  function end(v, why) {
    if (finished) { return; }
    finished = true;
    if (v === null) { try { x.abort(); } catch (e) {} }
    done(v, why);
  }
  try {
    x = new ActiveXObject("MSXML2.XMLHTTP.6.0");
    var u = authRawUrl(AUTH_URL);
    x.open("GET", u + (u.indexOf("?") < 0 ? "?" : "&") + "t=" + new Date().getTime(), true);
    x.setRequestHeader("Cache-Control", "no-cache");
    x.onreadystatechange = function () {
      if (x.readyState !== 4) { return; }
      if (x.status !== 200) { end(null, "인증 서버에 연결하지 못했습니다."); return; }
      var m = /^\s*([0-9a-f]{64})\s*$/i.exec(x.responseText);
      if (m) { end(m[1].toLowerCase()); } else { end(null, "인증 정보를 읽지 못했습니다(관리자 설정 확인 필요)."); }
    };
    x.send();
  } catch (e) { end(null, "인증 서버에 연결하지 못했습니다."); return; }
  window.setTimeout(function () { end(null, "인증 서버의 응답이 없습니다."); }, AUTH_TIMEOUT_MS);
}

// 창 전체를 덮고 인증을 확인한다. 통과하면 덮개를 걷고 done()을 부른다.
function authGate(appName, done) {
  authAppName = appName;
  if (!AUTH_URL) { usageDaily(); if (done) { done(); } return; }
  authLocked = true;
  authBuildOverlay(appName);
  authCheck(done);
}

function authCheck(done) {
  var st = authLoad();
  authShow("인증을 확인하는 중입니다...", false, false);
  authFetch(function (remote, why) {
    if (remote === null) {
      var age = new Date().getTime() - st.CheckedAt;
      if (st.Hash && age >= 0 && age <= AUTH_GRACE_DAYS * 86400000) { authUnlock(done); return; }
      authShow(why + "\n인터넷 연결을 확인한 뒤 [다시 시도]를 눌러 주세요.", false, true);
      authOnRetry = function () { authCheck(done); };
      return;
    }
    if (st.Hash === remote) { authSave(remote); authUnlock(done); return; }
    authShow(st.Hash ? "인증코드가 바뀌었습니다. 새 인증코드를 입력해 주세요." : "인증코드를 입력해 주세요.", true, false);
    authOnSubmit = function () {
      var el = document.getElementById("authCode");
      if (authHash(el.value) === remote) { authSave(remote); usageReport("인증"); authUnlock(done); return; }
      authShow("인증코드가 맞지 않습니다. 다시 입력해 주세요. (대소문자를 구분합니다)", true, false);
      el.value = "";
    };
  });
}

var authOnSubmit = null, authOnRetry = null;

function authBuildOverlay(appName) {
  var ov = document.createElement("div");
  ov.id = "authOverlay";
  ov.style.cssText = "position:absolute;top:0;left:0;right:0;bottom:0;z-index:10000;background:#F5F6F3;color:#1B231F;font-family:'Malgun Gothic','맑은 고딕',sans-serif;font-size:13px;";
  var btn = "height:30px;padding:0 14px;border:1px solid #00805D;border-radius:6px;background:#00805D;color:#FFFFFF;font-family:inherit;font-size:12.5px;font-weight:700;cursor:pointer;";
  ov.innerHTML =
    '<div style="position:absolute;top:50%;left:50%;width:380px;margin-left:-190px;margin-top:-105px;background:#FFFFFF;border:1px solid #C9D3CB;border-radius:8px;padding:20px 22px;">' +
      '<div id="authTitle" style="font-size:15px;font-weight:700;"></div>' +
      '<div id="authMsg" style="margin-top:10px;line-height:19px;min-height:38px;white-space:pre-line;"></div>' +
      '<div id="authForm" style="display:none;margin-top:8px;">' +
        '<input id="authCode" type="password" style="width:230px;height:30px;padding:3px 6px;border:1px solid #C9D3CB;border-radius:4px;font-family:inherit;font-size:13px;"> ' +
        '<button id="authOk" type="button" style="' + btn + '">확인</button>' +
      '</div>' +
      '<div id="authRetryRow" style="display:none;margin-top:8px;"><button id="authRetry" type="button" style="' + btn + '">다시 시도</button></div>' +
      '<div style="margin-top:14px;color:#5C685F;font-size:12px;">인증코드는 프로그램 관리자에게 문의하세요.</div>' +
    '</div>';
  document.body.appendChild(ov);
  document.getElementById("authTitle").innerText = appName + " - 인증";
  document.getElementById("authOk").onclick = function () { if (authOnSubmit) { authOnSubmit(); } };
  document.getElementById("authRetry").onclick = function () { if (authOnRetry) { authOnRetry(); } };
  document.getElementById("authCode").onkeydown = function (e) {
    e = e || window.event;
    if (e.keyCode === 13 && authOnSubmit) { authOnSubmit(); return false; }
  };
  // 덮개 밖(뒤에 있는 화면)으로 가는 클릭, 키 입력, 파일 끌어 놓기를 막는다(Tab으로 뒤쪽 버튼을 누르는 것까지)
  var block = function (e) {
    if (!authLocked) { return; }
    var inside = ov.contains(e.target);
    if (e.type === "drop" || e.type === "dragover" || e.type === "dragenter" || !inside) {
      e.preventDefault(); e.stopPropagation();
      if (!inside && e.type === "keydown") { try { document.getElementById("authCode").focus(); } catch (x) {} }
    }
  };
  var evs = ["click", "mousedown", "dblclick", "keydown", "keypress", "drop", "dragover", "dragenter"];
  for (var i = 0; i < evs.length; i++) { window.addEventListener(evs[i], block, true); }
}

function authShow(text, withForm, withRetry) {
  document.getElementById("authMsg").innerText = text;
  document.getElementById("authForm").style.display = withForm ? "block" : "none";
  document.getElementById("authRetryRow").style.display = withRetry ? "block" : "none";
  if (withForm) { try { document.getElementById("authCode").focus(); } catch (e) {} }
}

function authUnlock(done) {
  authLocked = false;
  var ov = document.getElementById("authOverlay");
  if (ov) { ov.parentNode.removeChild(ov); }
  usageDaily();
  if (done) { done(); }
}

// ---------------- 사용 기록 ----------------

// 프로그램을 연 기록은 PC마다 프로그램별로 하루 한 번만 보낸다. 보내는 데 성공한 날짜를 usage.json에 적는다
function usageDaily() {
  var today = usageToday(), sent = usageLoad();
  if (sent[authAppName] === today) { return; }
  usageReport("실행", function () { sent = usageLoad(); sent[authAppName] = today; usageSave(sent); });
}

// 설문지에 한 줄 보낸다. 기다리지 않고, 실패해도 프로그램 사용에는 영향이 없게 무시한다
var usageXhr = [];
function usageReport(kind, ok) {
  if (!USAGE_FORM_URL || !authAppName) { return; }
  try {
    var net = new ActiveXObject("WScript.Network");
    var body = USAGE_ENTRY.program + "=" + encodeURIComponent(authAppName) + "&" + USAGE_ENTRY.kind + "=" + encodeURIComponent(kind) +
      "&" + USAGE_ENTRY.pc + "=" + encodeURIComponent(net.ComputerName) + "&" + USAGE_ENTRY.user + "=" + encodeURIComponent(net.UserName);
    // XMLHTTP는 구글 설문지로 POST하면 '액세스 거부'가 난다(보안 영역 검사). ServerXMLHTTP는 그 검사를 하지 않는다
    var x = new ActiveXObject("MSXML2.ServerXMLHTTP.6.0");
    usageXhr.push(x);   // 응답이 오기 전에 정리되지 않게 잡아 둔다
    x.open("POST", USAGE_FORM_URL, true);
    x.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    x.onreadystatechange = function () {
      if (x.readyState !== 4) { return; }
      if (x.status === 200 && ok) { ok(); }
    };
    x.send(body);
  } catch (e) {}
}

function usageToday() {
  var d = new Date();
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
}
function usagePath() { return authStorePath().replace(/auth\.json$/, "usage.json"); }
function usageLoad() {
  try {
    var f = new ActiveXObject("Scripting.FileSystemObject").OpenTextFile(usagePath(), 1);
    var o = JSON.parse(f.ReadAll()); f.Close();
    return o && typeof o === "object" ? o : {};
  } catch (e) { return {}; }
}
function usageSave(o) {
  try {
    var fso = new ActiveXObject("Scripting.FileSystemObject"), p = usagePath(), dir = fso.GetParentFolderName(p);
    if (!fso.FolderExists(dir)) { fso.CreateFolder(dir); }
    var f = fso.CreateTextFile(p, true);
    f.Write(JSON.stringify(o)); f.Close();
  } catch (e) {}
}

// ---------------- SHA-256 (UTF-8) ----------------

function sha256Hex(text) {
  var K = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
  var H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  var s = unescape(encodeURIComponent(text)), bytes = [], i, t;
  for (i = 0; i < s.length; i++) { bytes.push(s.charCodeAt(i)); }
  var bitLen = bytes.length * 8;
  bytes.push(0x80);
  while (bytes.length % 64 !== 56) { bytes.push(0); }
  var hi = Math.floor(bitLen / 0x100000000), lo = bitLen >>> 0;
  bytes.push((hi >>> 24) & 255, (hi >>> 16) & 255, (hi >>> 8) & 255, hi & 255, (lo >>> 24) & 255, (lo >>> 16) & 255, (lo >>> 8) & 255, lo & 255);
  function rotr(x, n) { return (x >>> n) | (x << (32 - n)); }
  var W = new Array(64);
  for (var off = 0; off < bytes.length; off += 64) {
    for (t = 0; t < 16; t++) { W[t] = (bytes[off + t * 4] << 24) | (bytes[off + t * 4 + 1] << 16) | (bytes[off + t * 4 + 2] << 8) | bytes[off + t * 4 + 3]; }
    for (t = 16; t < 64; t++) {
      var x = W[t - 15], y = W[t - 2];
      var s0 = rotr(x, 7) ^ rotr(x, 18) ^ (x >>> 3), s1 = rotr(y, 17) ^ rotr(y, 19) ^ (y >>> 10);
      W[t] = (W[t - 16] + s0 + W[t - 7] + s1) | 0;
    }
    var a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
    for (t = 0; t < 64; t++) {
      var t1 = (h + (rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25)) + ((e & f) ^ (~e & g)) + K[t] + W[t]) | 0;
      var t2 = ((rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22)) + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
    H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
  }
  var hex = "";
  for (i = 0; i < 8; i++) { hex += ("00000000" + (H[i] >>> 0).toString(16)).slice(-8); }
  return hex;
}
