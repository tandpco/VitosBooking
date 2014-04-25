<script type="text/javascript">
<!--
// DISABLE F5 KEY
// Chrome
function my_onkeydown_handler() {
    switch (event.keyCode) {
        case 116 : // 'F5'
            event.preventDefault();
            event.keyCode = 0;
            window.status = "F5 disabled";
            break;
    }
}
document.addEventListener("keydown", my_onkeydown_handler);
// IE
document.onkeydown = function(){

if(window.event && window.event.keyCode == 116)
        { // Capture and remap F5
    window.event.keyCode = 505;
      }

if(window.event && window.event.keyCode == 505)
        { // New action for F5
    return false;
        // Must return false or the browser will refresh anyway
    }
}
//-->
</script>
<%
' 2011-05-27 TAM: Converted the following server side stuff to VBScript because we cannot mix.

'/*** Clock -- beginning of server-side support code
'by Andrew Shearer, http://www.shearersoftware.com/
'v2.1-ASP, 2002-10-12. For updates and explanations, see
'<http://www.shearersoftware.com/software/web-tools/clock/>. ***/
'
'/* Prevent this page from being cached (though some browsers still
'   cache the page anyway, which is why we use cookies). This is
'   only important if the cookie is deleted while the page is still
'   cached (and for ancient browsers that don't know about Cache-Control).
'   If that's not an issue, you may be able to get away with
'   "Cache-Control: private" instead. */
'Response.AddHeader("Pragma", "no-cache");
Response.AddHeader "Pragma", "no-cache"

'/* Grab the current server time. */
'var gDate = new Date();
Dim nOffset, gDate
If IsNumeric(Session("TZOffset")) Then
	nOffset = Session("TZOffset")
Else
	nOffset = 0
End If
gDate = DateAdd("h", nOffset, Now)

'/* Are the seconds shown by default? When changing this, also change the
'   JavaScript client code's definition of clockShowsSeconds below to match. */
'var clockShowsSeconds = false;
Dim clockShowsSeconds
clockShowsSeconds = FALSE

'function getServerDateItems() {
'    return gDate.getFullYear()+","+gDate.getMonth()+","+gDate.getDate()+","
'        +gDate.getHours()+","+gDate.getMinutes()+","+gDate.getSeconds();
'}
Function getServerDateItems()
	Dim sRet
	
	sRet = Year(gDate) & "," & (Month(gDate) - 1) & "," & Day(gDate) & "," & Hour(gDate) & "," & Minute(gDate) & "," & Second(gDate)
	
	getServerDateItems = sRet
End Function

'function clockDateString(inDate) {
'    return ["Sunday, ","Monday, ","Tuesday, ","Wednesday, ",
'            "Thursday, ","Friday, ","Saturday, "][inDate.getDay()]
'         + ["January ","February ","March ","April ","May ","June ",
'            "July ", "August ","September ","October ","November ","December "]
'            [inDate.getMonth()]
'         + inDate.getDate() + ", " + inDate.getFullYear();
'}
Function clockDateString(ByVal inDate)
	clockDateString = FormatDateTime(inDate, 1)
End Function

'function clockTimeString(inHours, inMinutes, inSeconds) {
'    return (inHours == 0 ? "12" : (inHours <= 12 ? inHours : inHours - 12))
'                + (inMinutes < 10 ? ":0" : ":") + inMinutes
'                + (clockShowsSeconds ? ((inSeconds < 10 ? ":0" : ":") + inSeconds) : "")
'                + (inHours < 12 ? " AM" : " PM");
'}
Function clockTimeString(ByVal inHours, inMinutes, inSeconds)
	Dim sRet
	
	If inHours = 0 Then
		sRet = "12:"
	Else
		If inHours > 12 Then
			sRet = (inHours - 12) & ":"
		Else
			sRet = inHours & ":"
		End If
	End If
	If inMinutes < 10 Then
		sRet = sRet & "0" & inMinutes
	Else
		sRet = sRet & inMinutes
	End If
	If clockShowsSeconds Then
		If inSeconds < 10 Then
			sRet = sRet & ":0" & inSeconds
		Else
			sRet = sRet & ":" & inSeconds
		End If
	End If
	If inHours < 12 Then
		sRet = sRet & " AM"
	Else
		sRet = sRet & " PM"
	End If
	
	clockTimeString = sRet
End Function

'/*** Clock -- end of server-side support code ***/
%>
<script language="JavaScript" type="text/javascript">
<!--
/* set up variables used to init clock in BODY's onLoad handler;
   should be done as early as possible */
var clockLocalStartTime = new Date();
var clockServerStartTime = new Date(<%=getServerDateItems()%>);

/* stub functions for older browsers;
   will be overridden by next JavaScript1.2 block */
function clockInit() {
}
//-->
</script>
<script language="JavaScript1.2" type="text/javascript">
<!--
/*** simpleFindObj, by Andrew Shearer

Efficiently finds an object by name/id, using whichever of the IE,
classic Netscape, or Netscape 6/W3C DOM methods is available.
The optional inLayer argument helps Netscape 4 find objects in
the named layer or floating DIV. */
function simpleFindObj(name, inLayer) {
	return document[name] || (document.all && document.all[name])
		|| (document.getElementById && document.getElementById(name))
		|| (document.layers && inLayer && document.layers[inLayer].document[name]);
}

/*** Beginning of Clock 2.0, by Andrew Shearer
See: http://www.shearersoftware.com/software/web-tools/clock/
Redistribution is permitted with the above notice intact.

Client-side clock, based on computed time differential between browser &
server. The server time is inserted by server-side JavaScript, and local
time is subtracted from it by client-side JavaScript while the page is
loading.

Cookies: The local and remote times are saved in cookies named
localClock and remoteClock, so that when the page is loaded from local
cache (e.g. by the Back button) the clock will know that the embedded
server time is stale compared to the local time, since it already
matches its cookie. It can then base the calculations on both cookies,
without reloading the page from the server. (IE 4 & 5 for Windows didn't
respect Response.Expires = 0, so if cookies weren't used, the clock
would be wrong after going to another page then clicking Back. Netscape
& Mac IE were OK.)

Every so often (by default, one hour) the clock will reload the page, to
make sure the clock is in sync (as well as to update the rest of the
page content).

Compatibility: IE 4.x and 5.0, Netscape 4.x and 6.0, Mozilla 1.0. Mac & Windows.

History:  1.0  5/9/2000 GIF-image digits
          2.0 6/29/2000 Uses text DIV layers (so 4.0 browsers req'd), &
                         cookies to work around Win IE stale-time bug
		  2.1 10/12/2002 Noted Mozilla 1.0 compatibility; released PHP version.
*/
var clockIncrementMillis = 60000;
var localTime;
var clockOffset;
var clockExpirationLocal;
var clockShowsSeconds = false;
var clockTimerID = null;

function clockInit(localDateObject, serverDateObject)
{
    var origRemoteClock = parseInt(clockGetCookieData("remoteClock"));
    var origLocalClock = parseInt(clockGetCookieData("localClock"));
    var newRemoteClock = serverDateObject.getTime();
    // May be stale (WinIE); will check against cookie later
    // Can't use the millisec. ctor here because of client inconsistencies.
    var newLocalClock = localDateObject.getTime();
    var maxClockAge = 60 * 60 * 1000;   // get new time from server every 1hr

    if (newRemoteClock != origRemoteClock) {
        // new clocks are up-to-date (newer than any cookies)
        document.cookie = "remoteClock=" + newRemoteClock;
        document.cookie = "localClock=" + newLocalClock;
        clockOffset = newRemoteClock - newLocalClock;
        clockExpirationLocal = newLocalClock + maxClockAge;
        localTime = newLocalClock;  // to keep clockUpdate() happy
    }
    else if (origLocalClock != origLocalClock) {
        // error; localClock cookie is invalid (parsed as NaN)
        clockOffset = null;
        clockExpirationLocal = null;
    }
    else {
        // fall back to clocks in cookies
        clockOffset = origRemoteClock - origLocalClock;
        clockExpirationLocal = origLocalClock + maxClockAge;
        localTime = origLocalClock;
        // so clockUpdate() will reload if newLocalClock
        // is earlier (clock was reset)
    }
    /* Reload page at server midnight to display the new date,
       by expiring the clock then */
    var nextDayLocal = (new Date(serverDateObject.getFullYear(),
            serverDateObject.getMonth(),
            serverDateObject.getDate() + 1)).getTime() - clockOffset;
    if (nextDayLocal < clockExpirationLocal) {
        clockExpirationLocal = nextDayLocal;
    }
}

function clockOnLoad()
{
    clockUpdate();
}

function clockOnUnload() {
    clockClearTimeout();
}

function clockClearTimeout() {
    if (clockTimerID) {
        clearTimeout(clockTimerID);
        clockTimerID = null;
    }
}

function clockToggleSeconds()
{
    clockClearTimeout();
    if (clockShowsSeconds) {
        clockShowsSeconds = false;
        clockIncrementMillis = 60000;
    }
    else {
        clockShowsSeconds = true;
        clockIncrementMillis = 1000;
    }
    clockUpdate();
}

function clockTimeString(inHours, inMinutes, inSeconds) {
    return inHours == null ? "-:--" : ((inHours == 0
                   ? "12" : (inHours <= 12 ? inHours : inHours - 12))
                + (inMinutes < 10 ? ":0" : ":") + inMinutes
                + (clockShowsSeconds
                   ? ((inSeconds < 10 ? ":0" : ":") + inSeconds) : "")
                + (inHours < 12 ? " AM" : " PM"));
}

function clockDisplayTime(inHours, inMinutes, inSeconds) {
    
    clockWriteToDiv("ClockTime", clockTimeString(inHours, inMinutes, inSeconds));
}

// 2011-05-27 TAM: Added routeines to display date too
function clockDateString(inDate) {
    return inDate == null ? "----, ---- --, ----" : (["Sunday, ","Monday, ","Tuesday, ","Wednesday, ",
            "Thursday, ","Friday, ","Saturday, "][inDate.getDay()]
         + ["January ","February ","March ","April ","May ","June ",
            "July ", "August ","September ","October ","November ","December "]
            [inDate.getMonth()]
         + inDate.getDate() + ", " + inDate.getFullYear());
}
function clockDisplayDate(inDate) {

	clockWriteToDiv("ClockDate", clockDateString(inDate));
}

function clockWriteToDiv(divName, newValue) // APS 6/29/00
{
    var divObject = simpleFindObj(divName);
    // 2011-05-27 TAM: Remove <p> tag
    //newValue = '<p>' + newValue + '<' + '/p>';
    if (divObject && divObject.innerHTML) {
        divObject.innerHTML = newValue;
    }
    else if (divObject && divObject.document) {
        divObject.document.writeln(newValue);
        divObject.document.close();
    }
    // else divObject wasn't found; it's only a clock, so don't bother complaining
}

function clockGetCookieData(label) {
    /* find the value of the specified cookie in the document's
    semicolon-delimited collection. For IE Win98 compatibility, search
    from the end of the string (to find most specific host/path) and
    don't require "=" between cookie name & empty cookie values. Returns
    null if cookie not found. One remaining problem: Under IE 5 [Win98],
    setting a cookie with no equals sign creates a cookie with no name,
    just data, which is indistinguishable from a cookie with that name
    but no data but can't be overwritten by any cookie with an equals
    sign. */
    var c = document.cookie;
    if (c) {
        var labelLen = label.length, cEnd = c.length;
        while (cEnd > 0) {
            var cStart = c.lastIndexOf(';',cEnd-1) + 1;
            /* bug fix to Danny Goodman's code: calculate cEnd, to
            prevent walking the string char-by-char & finding cookie
            labels that contained the desired label as suffixes */
            // skip leading spaces
            while (cStart < cEnd && c.charAt(cStart)==" ") cStart++;
            if (cStart + labelLen <= cEnd && c.substr(cStart,labelLen) == label) {
                if (cStart + labelLen == cEnd) {                
                    return ""; // empty cookie value, no "="
                }
                else if (c.charAt(cStart+labelLen) == "=") {
                    // has "=" after label
                    return unescape(c.substring(cStart + labelLen + 1,cEnd));
                }
            }
            cEnd = cStart - 1;  // skip semicolon
        }
    }
    return null;
}

/* Called regularly to update the clock display as well as onLoad (user
   may have clicked the Back button to arrive here, so the clock would need
   an immediate update) */
function clockUpdate()
{
    var lastLocalTime = localTime;
    localTime = (new Date()).getTime();
    
    /* Sanity-check the diff. in local time between successive calls;
       reload if user has reset system clock */
    if (clockOffset == null) {
        clockDisplayTime(null, null, null);
		// 2011-05-27 TAM: Added routeines to display date too
        clockDisplayDate(null);
    }
    else if (localTime < lastLocalTime || clockExpirationLocal < localTime) {
        /* Clock expired, or time appeared to go backward (user reset
           the clock). Reset cookies to prevent infinite reload loop if
           server doesn't give a new time. */
        document.cookie = 'remoteClock=-';
        document.cookie = 'localClock=-';
        location.reload();      // will refresh time values in cookies
    }
    else {
        // Compute what time would be on server 
        var serverTime = new Date(localTime + clockOffset);
        clockDisplayTime(serverTime.getHours(), serverTime.getMinutes(),
            serverTime.getSeconds());
		// 2011-05-27 TAM: Added routeines to display date too
        clockDisplayDate(serverTime);
        
        // Reschedule this func to run on next even clockIncrementMillis boundary
        clockTimerID = setTimeout("clockUpdate()",
            clockIncrementMillis - (serverTime.getTime() % clockIncrementMillis));
    }
}

/*** End of Clock ***/
//-->
</script>
