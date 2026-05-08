<!--#include file="./CCVerification_class.inc"-->

<HTML>
<HEAD>
<TITLE></TITLE>
<STYLE TYPE="text/css">
	h3{color:#CC3300;}
</STYLE>
<SCRIPT LANGUAGE="javascript">
function checkCCInput(ccForm) {
	var msg1   = new String("");
	var ccType = new Number(ccForm.cc_type.value);
	var ccNum  = new Number(ccForm.cc_num.value.length);
	var ccExp  = new Number(ccForm.cc_exp.value.length);
	if (ccType < 1 | ccType > 7 ) {
		msg1 += "The credit card type you have entered is invalid.";
		alert(msg1);
		ccForm.cc_type.focus();
		return(false);
	}
	if (ccNum < 13 | ccNum > 20) {
		msg1 += "The credit card number must be 13 to 20 ";
		msg1 += "characters in length, depending on the card\'s type.";
		alert(msg1);
		ccForm.cc_num.focus();
		return(false);
	}
	if (ccExp != 7) {
		msg1 += "The credit card date must be in \"mm\/yyyy\" format.";
		alert(msg1);
		ccForm.cc_exp.focus();
		return(false);
	}
	return(true);
}
</SCRIPT>
</HEAD>
<BODY BGCOLOR="#EEEEEE" BACKGROUND="/aspEmporium/bg.gif" TEXT="#000000">

<H3>CCVerification Object</h3>
<H3>Simple Credit Card Verification - Version 3.0</H3>


<%
dim cc, ccNum, ccType, ccExp

ccNum  = Request.Form("cc_num")    ' card number
ccType = Request.Form("cc_type")   ' card type
ccExp  = Request.Form("cc_exp")    ' card expiration date



 ' create an instance of the class

set cc = new CCVerification



 ' check user's input to make sure 
 ' something was entered for each input.

if _
    (Len(Trim(ccNum)) > 0)  And _
    (Len(Trim(ccType)) > 0) And _
    (Len(Trim(ccExp)) > 0) _
then
	with cc

		 ' set the properties of the CCVerification object

		.CardType           = CInt(  ccType )
		.CardNumber         = CStr(  ccNum  )
		.CardExpirationDate = CDate( ccExp  )



		 ' Verify the card with the VerifyCard method
		 ' and translate the results into a meaningful
		 ' message with the TranslateCardResults method.

		response.write( "Verification Results:" )
		response.write( "<BR><B>" & _
			.TranslateCardResults( .VerifyCard() ) & "</B><BR><BR>" )



		 ' base64 encode the card number and show the encoded number

		response.write( "Base64 Encoded Card Number:" )
		response.write( "<BR><B>" & _
			.EnDecryptCard(Base64Enc, .CardNumber) & "</B><BR><BR>" )



		 ' Un-comment the next line if you want to decode a base64 
		 ' encoded card number just replace the "THE CARDNUMBER TO DECODE" 
		 ' with a base64 encoded string.

		 ' response.write( .EnDecryptCard( Base64Dec, "THE CARDNUMBER TO DECODE" ) )
	end with
end if



 ' show the input form...

Response.Write( cc.ShowVerificationForm() )



 ' release the class from memory.

set cc = nothing
%>

</BODY>
</HTML>
