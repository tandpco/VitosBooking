<script type="text/javascript">

var ie4=document.all;

var gbHalfAddress = false;
var gbFocusAreaCode = false;

function resetRedirect() {
//  var loRedirectDiv;
//  
//  loRedirectDiv = ie4? eval("document.all.redirect") : document.getElementById("redirect");
//  loRedirectDiv.innerHTML = <%=gnRedirectTime%>;
}

function  showAllAddresses(el) {
  if(el.innerHTML == "All Addresses") {
    document.getElementById("addressList").className = 'showAll'
    el.innerHTML = 'Top 3 Only'
  } else {
    document.getElementById("addressList").className = ''
    el.innerHTML = 'All Addresses'
    $(el).closest('#content-wrapper').scrollTop('0')
  }
}


$(function(){

    $("#addressList button").each(function(){
      $(this).data('text',$(this).text())
    })
  $("#livesearch").on('change',function(){
    var $val = $(this).val()
    if($val) {
      document.getElementById("addressList").className = 'showAll'
      $("#toggleAddresssButtons").hide();
    }
    else {
      document.getElementById("addressList").className = ''
      $("#addressList button").removeClass('hidden')
      $("#content-wrapper").scrollTop('0')
      $("#toggleAddresssButtons").show();
    }
    console.log('changed',$val)
    $("#addressList button").each(function(){
      if($(this).data('text').indexOf($val) === -1) {
        $(this).html($(this).data('text'))
        $(this).addClass('hidden')
      }
      else {
        $(this).html($(this).data('text').replace($val, '<span class="highlight">'+$val+'</span>'))
        $(this).removeClass('hidden')
      }
    })

  })
})
function disableEnterKey() {
  var loText, loDiv;
  
  if (event.keyCode == 13) {
    event.cancelBubble = true;
    event.returnValue = false;
    return false;
  }
}

function getAddress() {
  var loText, loDiv;
  
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  loText.value = "";
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  loText.value = "";
  
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function addToPostalCode(psDigit) {
  var loText, lsText;
  
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  lsText = loText.value;
  lsText += psDigit;
  loText.value = lsText;
  
  resetRedirect();
}

function backspacePostalCode() {
  var loText, lsText;
  
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  lsText = loText.value;
  if (lsText.length > 0) {
    lsText = lsText.substr(0, (lsText.length - 1));
    loText.value = lsText;
  }
  
  resetRedirect();
}

function setPostalCode(psDigit) {
  var loPhone, lsPhone;
  
  loPhone = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  loPhone.value = psDigit;
  
  resetRedirect();
}

function addToStreetNumber(psDigit) {
  var loText, lsText;
  
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  lsText = loText.value;
  lsText += psDigit;
  loText.value = lsText;
  
  resetRedirect();
}

function backspaceStreetNumber() {
  var loText, lsText;
  
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  lsText = loText.value;
  if (lsText.length > 0) {
    lsText = lsText.substr(0, (lsText.length - 1));
    loText.value = lsText;
  }
  
  resetRedirect();
}

function cancelPostalCode() {
  var loText, loDiv;
  
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  loText.value = "";
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  loText.value = "";
  
  gbHalfAddress = false;
  
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "visible";
}

function getStreetLetter(pbHalfAddress) {
  var loText, lsPostalCode, lsStreetNumber, loDiv;
  
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  lsPostalCode = loText.value;
  if (lsPostalCode.length != 5)
    return false;
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  lsStreetNumber = loText.value;
  if (lsStreetNumber.length == 0)
    return false;
  
  gbHalfAddress = pbHalfAddress;
  
  loDiv = ie4? eval("document.all.postalstreetspan") : document.getElementById('postalstreetspan');
  loDiv.innerHTML = "<strong>Zip Code: " + lsPostalCode + " &nbsp; Street Number: " + lsStreetNumber
  if (gbHalfAddress) {
    loDiv.innerHTML = loDiv.innerHTML + " 1/2"
  }
  loDiv.innerHTML = loDiv.innerHTML + "</strong><br/>"
  
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function cancelAddress() {
  var loText, loDiv;
  
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function getName() {
  var loText, loDiv;
  
  loText = ie4? eval("document.all.name") : document.getElementById('name');
  loText.value = "";
  
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function addToName(psDigit) {
  var loName, lsName;
  
  loName = ie4? eval("document.all.name") : document.getElementById('name');
  lsName = loName.value;
  lsName += psDigit;
  loName.value = lsName;
  
  resetRedirect();
}

function backspaceName() {
  var loName, lsName;
  
  loName = ie4? eval("document.all.name") : document.getElementById('name');
  lsName = loName.value;
  if (lsName.length > 0) {
    lsName = lsName.substr(0, (lsName.length - 1));
    loName.value = lsName;
  }
  
  resetRedirect();
}

function cancelName() {
  var loDiv;
  
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function getNewPhone() {
  var loDiv;
  
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
}

function cancelNewPhone() {
<%
If gnOrderTypeID = 1 And ganAddressIDs(0) = 0 Then
%>
  var loAreaCode, loPhone, lsValue;
  
  loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
  if (loPhone.value.length != 8)
    return false;
  lsValue = loAreaCode.value + loPhone.value.substr(0, 3) + loPhone.value.substr(4);
  
  window.location = "/custmaint/newaddress.asp?t=<%=gnOrderTypeID%>&p=" + lsValue;
<%
Else
%>
  var loDiv;
  
  loDiv = ie4? eval("document.all.addressdiv") : document.getElementById('addressdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.postalcodediv") : document.getElementById('postalcodediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.namediv") : document.getElementById('namediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phonediv") : document.getElementById('phonediv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.phoneconfirmdiv") : document.getElementById('phoneconfirmdiv');
  loDiv.style.visibility = "hidden";
  loDiv = ie4? eval("document.all.assigndiv") : document.getElementById('assigndiv');
  loDiv.style.visibility = "visible";
  
  resetRedirect();
<%
End If
%>
}


function setFocusAreaCode(pbAreaCode) {
  var loAreaCode, loPhone;
  
  loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
  
  if (pbAreaCode) {
    loAreaCode.style.backgroundColor = "#FFFFFF";
    loPhone.style.backgroundColor = "#CCCCCC";
  }
  else {
    loAreaCode.style.backgroundColor = "#CCCCCC";
    loPhone.style.backgroundColor = "#FFFFFF";
  }
  
  gbFocusAreaCode = pbAreaCode;
  
  resetRedirect();
}

function addToPhone(psDigit) {
  var loPhone, lsPhone;
  
  if (gbFocusAreaCode) {
    loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  }
  else {
    loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
  }
  
  lsPhone = loPhone.value;
  if (gbFocusAreaCode) {
    if (lsPhone.length < 3) {
      lsPhone += psDigit;
      loPhone.value = lsPhone;
    }
    if (lsPhone.length == 3) {
      setFocusAreaCode(false);
    }
  }
  else {
    if (lsPhone.length < 8) {
      if (lsPhone.length == 3) {
        lsPhone = lsPhone + "-";
      }
      lsPhone += psDigit;
      loPhone.value = lsPhone;
    }
  }
  
  resetRedirect();
}

function setAreaCode(psDigit) {
  var loPhone, lsPhone;
  
  loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  lsPhone = psDigit;
  loPhone.value = lsPhone;
  
  resetRedirect();
}

function clearAreaCode(psDigit) {
  var loPhone, lsPhone;
  
  loPhone = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  loPhone.value = "";
  
  setFocusAreaCode(true);
}

function backspacePhone() {
  var loText, lsText;
  
  if (gbFocusAreaCode) {
    loText = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  }
  else {
    loText = ie4? eval("document.all.phone") : document.getElementById('phone');
  }
  
  lsText = loText.value;
  if (lsText.length > 0) {
    lsText = lsText.substr(0, (lsText.length - 1));
    if ((!gbFocusAreaCode) && (lsText.length == 4)) {
      lsText = lsText.substr(0, (lsText.length - 1));
    }
    loText.value = lsText;
  }
  
  resetRedirect();
}

function cancelPhone() {
  var loPhone, loOrderTypeDiv, loPhoneDiv, loNameDiv;
  
  window.location = "neworder.asp";
}

function goNewPhone() {
  var loName, lsName, loAreaCode, loPhone, lsValue, lsLocation;
  
  loAreaCode = ie4? eval("document.all.areacode") : document.getElementById('areacode');
  loPhone = ie4? eval("document.all.phone") : document.getElementById('phone');
  if (loAreaCode.value.length != 3)
    return false;
  if (loPhone.value.length != 8)
    return false;
  lsValue = loAreaCode.value + loPhone.value.substr(0, 3) + loPhone.value.substr(4);
  
  lsLocation = "customerfind.asp?t=<%=gnOrderTypeID%>&p=" + lsValue + "&p2=yes";
  
  window.location = lsLocation;
}

function goNext(psDigit) {
  var loText, lsLocation;
  
  lsLocation = "streetfind.asp?t=<%=Request("t")%>&z=";
  loText = ie4? eval("document.all.postalcode") : document.getElementById('postalcode');
  lsLocation = lsLocation + encodeURIComponent(loText.value) + "&y=";
  
  loText = ie4? eval("document.all.streetnumber") : document.getElementById('streetnumber');
  lsLocation = lsLocation + encodeURIComponent(loText.value) + "&x=" + psDigit;
  
  if (gbHalfAddress) {
    lsLocation = lsLocation + "&w=true";
  }
  else {
    lsLocation = lsLocation + "&w=false";
  }
  lsLocation = lsLocation + "&c=" + gnCustomerID.toString();
  
  window.location = lsLocation;
}

function goPickupNoCustomer() {
  var loText, lsLocation;
  
  loText = ie4? eval("document.all.name") : document.getElementById('name');
  lsText = loText.value;
  if (lsText.length == 0) {
    return false;
  }
  
  lsLocation = "unitselect.asp?t=<%=gnOrderTypeID%>&n=" + encodeURIComponent(lsText);
  
  window.location = lsLocation;
}

function verifyClick() {
    alert("Hello this is an Alert");
}

function back2Delivery() {
    var lsLocation = "neworder.asp";
//    alert("Back 2 Delivery");
    window.location = lsLocation;
}

function back2Phone() {
    var lsLocation = "neworder.asp";
//    alert("Back 2 Phone");
    window.location = lsLocation;
}


//-->
</script>