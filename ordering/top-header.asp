<script type="text/javascript">
  function changeDeliveryType() {
    var $s = window.location.search;
    if($s.indexOf('t=1') !== -1) {
      $s = $s.replace('t=1','t=2')
    }

    else if($s.indexOf('t=2') !== -1) {
      $s = $s.replace('t=2','t=1')
    }
    window.location = window.location.pathname+$s
  }
</script>
<tr height="72">
  <td valign="top" width="1010" height="72">
    <div id="statusBlock">
      <strong><%=IIf(gbTestMode,IIf(gbDevMode,"[DEV]","[TEST]"),"")%> Store <%=Session("StoreID")%></strong> |
      <b><%=Session("name")%></b>
      <div>
      <span id="ClockDate"><%=clockDateString(gDate)%></span> |
      <span id="ClockTime" onclick="clockToggleSeconds()"><%=clockTimeString(Hour(gDate), Minute(gDate), Second(gDate))%></span></div>
    </div>
    <ol id="tabs">
      <li class="<%=IIf(currentTab = "start","active","")%>"><a onclick="if(confirm('Would you like to toggle pickup/delivery?')) changeDeliveryType();" title="Delivery" class="showExtraNoteOnHover"><%= Iif(Session("OrderTypeID") = 2,"Pickup", "Delivery")%></a></li>
      <li class="<%=IIf(currentTab = "phone","active","")%>"><a onclick="if(confirm('Are you sure? You will lose any data you have entered so far.')) back2Phone();" title="Phone">Phone</a></li>
      <li class="<%=IIf(currentTab = "address","active","")%>">
        <% If Session("CustomerPhone") Then %>
          <a href="/ordering/customerfind.asp?t=<%=Session("OrderTypeID")%>&amp;p=<%=Session("CustomerPhone")%>">Address</a>
        <% Else %>
          <a href="/ordering/customerfind.php?t=<">Address</a>
        <% End If %>
      </li>
      <li class="<%=IIf(Session("AddressID") <= 1,"disabled","")%> <%=IIf(currentTab = "customer-name","active","")%>">
        <% If Session("AddressID") > 1 Then %>
          <a href="/ordering/customerselect.asp?t=<%=Session("OrderTypeID")%>&amp;a=<%=Session("AddressID")%>">Customer Name</a>
        <% Else %>
        Customer Name
        <% End If %>
      </li>
      <li class="<%=IIf(Session("AddressID") <= 1,"disabled","")%>  <%=IIf(currentTab = "order","active","")%>">Order</li>
      <li class="<%=IIf(Session("AddressID") <= 1,"disabled","")%>  <%=IIf(currentTab = "notes","active","")%>" onclick="gotoOrderNotes()">Notes</li>
    </ol>
  </td>
</tr>