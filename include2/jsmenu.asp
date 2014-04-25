<script type="text/javascript">
<!--
var i, j, k;

var gdOrderTotal = <%=gdOrderTotal%>;
var gnOrderLineID = <%=gnOrderLineID%>;
var gnUnitID = <%=gnUnitID%>;
var gnSpecialtyID = <%=gnSpecialtyID%>;
var gnSizeID = <%=gnSizeID%>;
var gnStyleID = <%=gnStyleID%>;
var gnHalf1SauceID = <%=gnHalf1SauceID%>;
var gnHalf2SauceID = <%=gnHalf2SauceID%>;
var gnHalf1SauceModifierID = <%=gnHalf1SauceModifierID%>;
var gnHalf2SauceModifierID = <%=gnHalf2SauceModifierID%>;
var gsOrderLineNotes = "<%=gsOrderLineNotes%>";
var gdOrderLineCost = <%=gdOrderLineCost%>;
var gdOrderLineDiscount = <%=gdOrderLineDiscount%>;
var gnCouponID = <%=gnCouponID%>;
var gsUnitDescription = "<%=gsUnitDescription%>";
var gsSpecialtyDescription = "<%=gsSpecialtyDescription%>";
var gsSizeDescription = "<%=gsSizeDescription%>";
var gsStyleDescription = "<%=gsStyleDescription%>";
var gsHalf1SauceDescription = "<%=gsHalf1SauceDescription%>";
var gsHalf2SauceDescription = "<%=gsHalf2SauceDescription%>";
var gsHalf1SauceModifierDescription = "<%=gsHalf1SauceModifierDescription%>";
var gsHalf2SauceModifierDescription = "<%=gsHalf2SauceModifierDescription%>";
var gdQuantity = 1;

i = <%=UBound(ganItemIDs) + 1%>;
var ganItemIDs = new Array(i);
var ganHalfIDs = new Array(i);
var gasItemDescriptions = new Array(i);
var gasItemShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganItemIDs)
%>
	ganItemIDs[<%=i%>] = <%=ganItemIDs(i)%>;
	ganHalfIDs[<%=i%>] = <%=ganHalfIDs(i)%>;
	gasItemDescriptions[<%=i%>] = "<%=gasItemDescriptions(i)%>";
	gasItemShortDescriptions[<%=i%>] = "<%=gasItemShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganTopperIDs) + 1%>;
var ganTopperIDs = new Array(i);
var ganTopperHalfIDs = new Array(i);
var gasTopperDescriptions = new Array(i);
var gasTopperShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganTopperIDs)
%>
	ganTopperIDs[<%=i%>] = <%=ganTopperIDs(i)%>;
	ganTopperHalfIDs[<%=i%>] = <%=ganTopperHalfIDs(i)%>;
	gasTopperDescriptions[<%=i%>] = "<%=gasTopperDescriptions(i)%>";
	gasTopperShortDescriptions[<%=i%>] = "<%=gasTopperShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganFreeSideIDs) + 1%>;
var ganFreeSideIDs = new Array(i);
var gasFreeSideDescriptions = new Array(i);
var gasFreeSideShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganFreeSideIDs)
%>
	ganFreeSideIDs[<%=i%>] = <%=ganFreeSideIDs(i)%>;
	gasFreeSideDescriptions[<%=i%>] = "<%=gasFreeSideDescriptions(i)%>";
	gasFreeSideShortDescriptions[<%=i%>] = "<%=gasFreeSideShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganAddSideIDs) + 1%>;
var ganAddSideIDs = new Array(i);
var gasAddSideDescriptions = new Array(i);
var gasAddSideShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganAddSideIDs)
%>
	ganAddSideIDs[<%=i%>] = <%=ganAddSideIDs(i)%>;
	gasAddSideDescriptions[<%=i%>] = "<%=gasAddSideDescriptions(i)%>";
	gasAddSideShortDescriptions[<%=i%>] = "<%=gasAddSideShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganUnitSizeIDs) + 1%>;
var ganUnitSizeIDs = new Array(i);
var gasUnitSizeDescriptions = new Array(i);
var gasUnitSizeShortDescriptions = new Array(i);
var gadUnitSizeStandardBasePrice = new Array(i);
var ganUnitSizeStandardNumberIncludedItems = new Array(i);
var gadUnitSizeSpecialtyBasePrice = new Array(i);
var ganUnitSizeSpecialtyNumberIncludedItems = new Array(i);
var ganUnitSizePercentSpecialtyItemVariances = new Array(i);
var gadPerAdditionalItemPrices = new Array(i);
var gabIsTaxable = new Array(i);

<%
For i = 0 To UBound(ganUnitSizeIDs)
%>
	ganUnitSizeIDs[<%=i%>] = <%=ganUnitSizeIDs(i)%>;
	gasUnitSizeDescriptions[<%=i%>] = "<%=gasUnitSizeDescriptions(i)%>";
	gasUnitSizeShortDescriptions[<%=i%>] = "<%=gasUnitSizeShortDescriptions(i)%>";
	gadUnitSizeStandardBasePrice[<%=i%>] = <%=gadUnitSizeStandardBasePrice(i)%>;
	ganUnitSizeStandardNumberIncludedItems[<%=i%>] = <%=ganUnitSizeStandardNumberIncludedItems(i)%>;
	gadUnitSizeSpecialtyBasePrice[<%=i%>] = <%=gadUnitSizeSpecialtyBasePrice(i)%>;
	ganUnitSizeSpecialtyNumberIncludedItems[<%=i%>] = <%=ganUnitSizeSpecialtyNumberIncludedItems(i)%>;
	ganUnitSizePercentSpecialtyItemVariances[<%=i%>] = <%=ganUnitSizePercentSpecialtyItemVariances(i)%>;
	gadPerAdditionalItemPrices[<%=i%>] = <%=gadPerAdditionalItemPrices(i)%>;
	gabIsTaxable[<%=i%>] = <%=LCase(gabIsTaxable(i))%>
<%
Next
%>

i = <%=UBound(ganSizeStyleIDs) + 1%>;
j = <%=UBound(ganSizeStyleSizeIDs, 2) + 1%>;
var ganSizeStyleIDs = new Array(i);
var gasSizeStyleDescriptions = new Array(i);
var gasSizeStyleShortDescriptions = new Array(i);
var gasSizeStyleSpecialMessage = new Array(i);
var ganSizeStyleSizeIDs = new Array(i);
var gadSizeStyleSurcharges = new Array(i);
<%
For i = 0 To UBound(ganSizeStyleIDs)
%>
	ganSizeStyleSizeIDs[<%=i%>] = new Array(j);
	gadSizeStyleSurcharges[<%=i%>] = new Array(j);
<%
Next
%>

<%
For i = 0 To UBound(ganSizeStyleIDs)
%>
	ganSizeStyleIDs[<%=i%>] = <%=ganSizeStyleIDs(i)%>;
	gasSizeStyleDescriptions[<%=i%>] = "<%=gasSizeStyleDescriptions(i)%>";
	gasSizeStyleShortDescriptions[<%=i%>] = "<%=gasSizeStyleShortDescriptions(i)%>";
	gasSizeStyleSpecialMessage[<%=i%>] = "<%=Server.HTMLEncode(gasSizeStyleSpecialMessage(i))%>";
<%
	For j = 0 To UBound(ganSizeStyleSizeIDs, 2)
%>
		ganSizeStyleSizeIDs[<%=i%>][<%=j%>] = <%=ganSizeStyleSizeIDs(i, j)%>;
		gadSizeStyleSurcharges[<%=i%>][<%=j%>] = <%=gadSizeStyleSurcharges(i, j)%>;
<%
	Next
Next
%>

i = <%=UBound(ganSauceIDs) + 1%>;
var ganSauceIDs = new Array(i);
var gasSauceDescriptions = new Array(i);
var gasSauceShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganSauceIDs)
%>
	ganSauceIDs[<%=i%>] = <%=ganSauceIDs(i)%>;
	gasSauceDescriptions[<%=i%>] = "<%=gasSauceDescriptions(i)%>";
	gasSauceShortDescriptions[<%=i%>] = "<%=gasSauceShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganSauceModifierIDs) + 1%>;
var ganSauceModifierIDs = new Array(i);
var gasSauceModifierDescriptions = new Array(i);
var gasSauceModifierShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganSauceModifierIDs)
%>
	ganSauceModifierIDs[<%=i%>] = <%=ganSauceModifierIDs(i)%>;
	gasSauceModifierDescriptions[<%=i%>] = "<%=gasSauceModifierDescriptions(i)%>";
	gasSauceModifierShortDescriptions[<%=i%>] = "<%=gasSauceModifierShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganUnitItemIDs) + 1%>;
var ganUnitItemIDs = new Array(i);
var gasUnitItemDescriptions = new Array(i);
var gasUnitItemShortDescriptions = new Array(i);
var gadUnitItemOnSidePrice = new Array(i);
var ganUnitItemCounts = new Array(i);
var gabUnitFreeItemFlags = new Array(i);
var gabUnitItemIsCheeses = new Array(i);
var gabUnitItemIsBaseCheeses = new Array(i);
var gabUnitItemIsExtraCheeses = new Array(i);

<%
For i = 0 To UBound(ganUnitItemIDs)
%>
	ganUnitItemIDs[<%=i%>] = <%=ganUnitItemIDs(i)%>;
	gasUnitItemDescriptions[<%=i%>] = "<%=gasUnitItemDescriptions(i)%>";
	gasUnitItemShortDescriptions[<%=i%>] = "<%=gasUnitItemShortDescriptions(i)%>";
	gadUnitItemOnSidePrice[<%=i%>] = <%=gadUnitItemOnSidePrice(i)%>;
	ganUnitItemCounts[<%=i%>] = <%=ganUnitItemCounts(i)%>;
	gabUnitFreeItemFlags[<%=i%>] = <%=LCase(gabUnitFreeItemFlags(i))%>;
	gabUnitItemIsCheeses[<%=i%>] = <%=LCase(gabUnitItemIsCheeses(i))%>;
	gabUnitItemIsBaseCheeses[<%=i%>] = <%=LCase(gabUnitItemIsBaseCheeses(i))%>;
	gabUnitItemIsExtraCheeses[<%=i%>] = <%=LCase(gabUnitItemIsExtraCheeses(i))%>;
<%
Next
%>

i = <%=UBound(ganUnitTopperIDs) + 1%>;
var ganUnitTopperIDs = new Array(i);
var gasUnitTopperDescriptions = new Array(i);
var gasUnitTopperShortDescriptions = new Array(i);

<%
For i = 0 To UBound(ganUnitTopperIDs)
%>
	ganUnitTopperIDs[<%=i%>] = <%=ganUnitTopperIDs(i)%>;
	gasUnitTopperDescriptions[<%=i%>] = "<%=gasUnitTopperDescriptions(i)%>";
	gasUnitTopperShortDescriptions[<%=i%>] = "<%=gasUnitTopperShortDescriptions(i)%>";
<%
Next
%>

i = <%=UBound(ganUnitSideIDs) + 1%>;
var ganUnitSideIDs = new Array(i);
var gasUnitSideDescriptions = new Array(i);
var gasUnitSideShortDescriptions = new Array(i);
var gadUnitSidePrices = new Array(i);

<%
For i = 0 To UBound(ganUnitSideIDs)
%>
	ganUnitSideIDs[<%=i%>] = <%=ganUnitSideIDs(i)%>;
	gasUnitSideDescriptions[<%=i%>] = "<%=gasUnitSideDescriptions(i)%>";
	gasUnitSideShortDescriptions[<%=i%>] = "<%=gasUnitSideShortDescriptions(i)%>";
	gadUnitSidePrices[<%=i%>] = <%=gadUnitSidePrices(i)%>;
<%
Next
%>

i = <%=UBound(ganUnitSpecialtyIDs) + 1%>;
j = <%=UBound(ganUnitSpecialtyItemIDs, 2) + 1%>;
var ganUnitSpecialtyIDs = new Array(i);
var gasUnitSpecialtyDescriptions = new Array(i);
var gasUnitSpecialtyShortDescriptions = new Array(i);
var ganUnitSpecialtySauceID = new Array(i);
var ganUnitSpecialtyStyleID = new Array(i);
var gabSpecialtyNoBaseCheese = new Array(i);
var ganUnitSpecialtyItemIDs = new Array(i);
var ganUnitSpecialtyItemQuantity = new Array(i);
<%
For i = 0 To UBound(ganUnitSpecialtyIDs)
%>
	ganUnitSpecialtyItemIDs[<%=i%>] = new Array(j);
	ganUnitSpecialtyItemQuantity[<%=i%>] = new Array(j);
<%
Next
%>

<%
For i = 0 To UBound(ganUnitSpecialtyIDs)
%>
	ganUnitSpecialtyIDs[<%=i%>] = <%=ganUnitSpecialtyIDs(i)%>;
	gasUnitSpecialtyDescriptions[<%=i%>] = "<%=gasUnitSpecialtyDescriptions(i)%>";
	gasUnitSpecialtyShortDescriptions[<%=i%>] = "<%=gasUnitSpecialtyShortDescriptions(i)%>";
	ganUnitSpecialtySauceID[<%=i%>] = <%=ganUnitSpecialtySauceID(i)%>;
	ganUnitSpecialtyStyleID[<%=i%>] = <%=ganUnitSpecialtyStyleID(i)%>;
	gabSpecialtyNoBaseCheese[<%=i%>] = <%=LCase(gabSpecialtyNoBaseCheese(i))%>;
<%
	For j = 0 To UBound(ganUnitSpecialtyItemIDs, 2)
%>
		ganUnitSpecialtyItemIDs[<%=i%>][<%=j%>] = <%=ganUnitSpecialtyItemIDs(i, j)%>;
		ganUnitSpecialtyItemQuantity[<%=i%>][<%=j%>] = <%=ganUnitSpecialtyItemQuantity(i, j)%>;
<%
	Next
Next
%>

i = <%=UBound(ganUpchargeSizeIDs) + 1%>;
j = <%=UBound(ganUpchargeItemIDs, 2) + 1%>;
var ganUpchargeSizeIDs = new Array(i);
var ganUpchargeItemIDs = new Array(i);
var gadUpchargePrice = new Array(i);
<%
For i = 0 To UBound(ganUpchargeSizeIDs)
%>
	ganUpchargeItemIDs[<%=i%>] = new Array(j);
	gadUpchargePrice[<%=i%>] = new Array(j);
<%
Next
%>

<%
For i = 0 To UBound(ganUpchargeSizeIDs)
%>
	ganUpchargeSizeIDs[<%=i%>] = <%=ganUpchargeSizeIDs(i)%>;
<%
	For j = 0 To UBound(ganUpchargeItemIDs, 2)
%>
		ganUpchargeItemIDs[<%=i%>][<%=j%>] = <%=ganUpchargeItemIDs(i, j)%>;
		gadUpchargePrice[<%=i%>][<%=j%>] = <%=gadUpchargePrice(i, j)%>;
<%
	Next
Next
%>

i = <%=UBound(ganSideGroupSpecialtyIDs) + 1%>;
j = <%=UBound(ganSideGroupSizeIDs, 2) + 1%>;
k = <%=UBound(ganSideGroupSideGroupIDs, 3) + 1%>;
var ganSideGroupSpecialtyIDs = new Array(i);
var ganSideGroupSizeIDs = new Array(i);
var ganSideGroupSideGroupIDs = new Array(i);
var gadSideGroupQuantity = new Array(i);
<%
For i = 0 To UBound(ganSideGroupSpecialtyIDs)
%>
	ganSideGroupSizeIDs[<%=i%>] = new Array(j);
	ganSideGroupSideGroupIDs[<%=i%>] = new Array(j);
	gadSideGroupQuantity[<%=i%>] = new Array(j);
<%
Next
For i = 0 To UBound(ganSideGroupSpecialtyIDs)
	For j = 0 To UBound(ganSideGroupSizeIDs, 2)
%>
	ganSideGroupSideGroupIDs[<%=i%>][<%=j%>] = new Array(k);
	gadSideGroupQuantity[<%=i%>][<%=j%>] = new Array(k);
<%
	Next
Next
%>

<%
For i = 0 To UBound(ganSideGroupSpecialtyIDs)
%>
	ganSideGroupSpecialtyIDs[<%=i%>] = <%=ganSideGroupSpecialtyIDs(i)%>;
<%
	For j = 0 To UBound(ganSideGroupSizeIDs, 2)
%>
		ganSideGroupSizeIDs[<%=i%>][<%=j%>] = <%=ganSideGroupSizeIDs(i, j)%>;
<%
		For k = 0 To UBound(ganSideGroupSideGroupIDs, 3)
%>
			ganSideGroupSideGroupIDs[<%=i%>][<%=j%>][<%=k%>] = <%=ganSideGroupSideGroupIDs(i, j, k)%>;
			gadSideGroupQuantity[<%=i%>][<%=j%>][<%=k%>] = <%=gadSideGroupQuantity(i, j, k)%>;
<%
		Next
	Next
Next
%>

i = <%=UBound(ganUnitGroupSizeIDs) + 1%>;
j = <%=UBound(ganUnitGroupSideGroupIDs, 2) + 1%>;
var ganUnitGroupSizeIDs = new Array(i);
var ganUnitGroupSideGroupIDs = new Array(i);
var gadUnitGroupQuantity = new Array(i);
<%
For i = 0 To UBound(ganUnitGroupSizeIDs)
%>
	ganUnitGroupSideGroupIDs[<%=i%>] = new Array(j);
	gadUnitGroupQuantity[<%=i%>] = new Array(j);
<%
Next
%>

<%
For i = 0 To UBound(ganUnitGroupSizeIDs)
%>
	ganUnitGroupSizeIDs[<%=i%>] = <%=ganUnitGroupSizeIDs(i)%>;
<%
	For j = 0 To UBound(ganUnitGroupSideGroupIDs, 2)
%>
		ganUnitGroupSideGroupIDs[<%=i%>][<%=j%>] = <%=ganUnitGroupSideGroupIDs(i, j)%>;
		gadUnitGroupQuantity[<%=i%>][<%=j%>] = <%=gadUnitGroupQuantity(i, j)%>;
<%
	Next
Next
%>

i = <%=UBound(ganSideGroupIDs) + 1%>;
j = <%=UBound(ganSideGroupSideIDs, 2) + 1%>;
var ganSideGroupIDs = new Array(i);
var gasSideGroupDescriptions = new Array(i);
var gasSideGroupShortDescriptions = new Array(i);
var ganSideGroupSideIDs = new Array(i);
var gasSideGroupSideDescriptions = new Array(i);
var gasSideGroupSideShortDescriptions = new Array(i);
var gabSideGroupSideIsDefault = new Array(i);
<%
For i = 0 To UBound(ganSideGroupSideIDs)
%>
	ganSideGroupSideIDs[<%=i%>] = new Array(j);
	gasSideGroupSideDescriptions[<%=i%>] = new Array(j);
	gasSideGroupSideShortDescriptions[<%=i%>] = new Array(j);
	gabSideGroupSideIsDefault[<%=i%>] = new Array(j);
<%
Next
%>

<%
For i = 0 To UBound(ganSideGroupIDs)
%>
	ganSideGroupIDs[<%=i%>] = <%=ganSideGroupIDs(i)%>;
	gasSideGroupDescriptions[<%=i%>] = "<%=gasSideGroupDescriptions(i)%>";
	gasSideGroupShortDescriptions[<%=i%>] = "<%=gasSideGroupShortDescriptions(i)%>";
<%
	For j = 0 To UBound(ganSideGroupSideIDs, 2)
%>
		ganSideGroupSideIDs[<%=i%>][<%=j%>] = <%=ganSideGroupSideIDs(i, j)%>;
		gasSideGroupSideDescriptions[<%=i%>][<%=j%>] = "<%=gasSideGroupSideDescriptions(i, j)%>";
		gasSideGroupSideShortDescriptions[<%=i%>][<%=j%>] = "<%=gasSideGroupSideShortDescriptions(i, j)%>";
		gabSideGroupSideIsDefault[<%=i%>][<%=j%>] = <%=LCase(gabSideGroupSideIsDefault(i, j))%>;
<%
	Next
Next
%>

var gbHasSpecSides = false;
if (gnSpecialtyID != 0) {
	for (i = 0; i < ganSideGroupSpecialtyIDs.length; i++) {
		if (ganSideGroupSpecialtyIDs[i] == gnSpecialtyID) {
			for (j = 0; j < ganSideGroupSizeIDs[i].length; j++) {
				if (ganSideGroupSizeIDs[i][j] == gnSizeID) {
					if (ganSideGroupSideGroupIDs[i][j][0] != 0) {
						gbHasSpecSides = true;
					}
			
					j = ganSideGroupSizeIDs[i].length;
				}
			}
			
			i = ganSideGroupSpecialtyIDs.length;
		}
	}
}

function FormatCurrency(amount)
{
	var i = parseFloat(amount);
	if(isNaN(i)) { i = 0.00; }
	var minus = '';
	if(i < 0) { minus = '-'; }
	i = Math.abs(i);
	i = parseInt((i + .005) * 100);
	i = i / 100;
	s = new String(i);
	if(s.indexOf('.') < 0) { s += '.00'; }
	if(s.indexOf('.') == (s.length - 2)) { s += '0'; }
	s = '$' + minus + s;
	return s;
}

function recalculatePrice() {
	var i, j, k, l, ldPrice, lnIncludedItems, lnTmp, lnItemVariance, ldPerItemPrices, lnItemCount, ldPremiumSurcharge, ldOrderPrice, loTotalDiv, lbIgnoreSpecialty;
	
	ldPrice = 0.00;
	
	// REMOVED - NOT NECESSARY FOR DEV
	
	s = "This Unit: " + FormatCurrency(gdOrderLineCost) + "&nbsp;&nbsp;&nbsp; Total: " + FormatCurrency(ldOrderPrice);
	loTotalDiv = ie4? eval("document.all.totaldiv") : document.getElementById("totaldiv");
	loTotalDiv.innerHTML = s;
}
//-->
</script>