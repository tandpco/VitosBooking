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
      <li class="<%=IIf(currentTab = "start","active","")%>"><a onclick="back2Delivery();" title="Delivery">Delivery</a></li>
      <li class="<%=IIf(currentTab = "phone","active","")%>"><a onclick="back2Phone();" title="Phone">Phone</a></li>
      <li class="<%=IIf(currentTab = "address","active","")%>">
        <% If Session("CustomerPhone") Then %>
          <a href="customerfind.asp?t=<%=gnOrderTypeID%>&amp;p=<%=Session("CustomerPhone")%>">Address</a>
        <% Else %>
          <a href="customerfind.php?t=<">Address</a>
        <% End If %>
      </li>
      <li class="<%=IIf(currentTab = "customer-name","active","")%>">
        <% If Session("AddressID") Then %>
          <a href="customerselect.asp?t=<%=gnOrderTypeID%>&amp;a=<%=Session("AddressID")%>">Customer Name</a>
        <% Else %>
          <a href="customerselect.php">Customer Name</a>
        <% End If %>
      </li>
      <li class="<%=IIf(currentTab = "order","active","")%>">Order</li>
      <li class="<%=IIf(currentTab = "notes","active","")%>">Notes</li>
    </ol>
  </td>
</tr>