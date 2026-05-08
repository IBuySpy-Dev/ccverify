<%
 ' #################################
 ' #  CCVerification Object  v3.0  #
 ' #  http://www.aspEmporium.com/  #
 ' #                               ##############################################
 ' #                                                                            #
 ' # PURPOSE:                                                                   #
 ' #  This class verifies the accuracy of an entered credit card number.        #
 ' #  IT DOES NOT VALIDATE OR IN ANY WAY PROCESS OR STORE A CREDIT CARD NUMBER! #
 ' #  What it does do is check the prefix and length of a card number           #
 ' #  against strict standards published by the credit card companies           #
 ' #  themselves. It also checks card numbers using the Luhn Check Digit        #
 ' #  Algorithm and verifies the expiration date and card type. Credit card     #
 ' #  information can be Base64 encoded for extra security if necessary.        #
 ' #                                                                            #
 ' #                                                                            #
 ' # VERSION NOTES:                                                             #
 ' #  +-------------+                                                           #
 ' #  | VERSION 1.0 |                                                           #
 ' #  |             +--------------------------------------------------------+  #
 ' #  |   created on: 5/2/2000                                               |  #
 ' #  |    - credit card verification implemented as a single function       |  #
 ' #  |      called "VerifyCreditCardNumber" with a few bugs                 |  #
 ' #  +----------------------------------------------------------------------+  #
 ' #                                                                            #
 ' #  +-------------+                                                           #
 ' #  | VERSION 2.0 |                                                           #
 ' #  |             +--------------------------------------------------------+  #
 ' #  |   released on: 8/17/2000                                             |  #
 ' #  |    - "VerifyCreditCardNumber" function modified by Tarun Bhushan to  |  #
 ' #  |      include the following features:                                 |  #
 ' #  |        = Added Luhn Check Digit Routine                              |  #
 ' #  |        = Added return values and verbose strings                     |  #
 ' #  |        = Fixed 2 to 4 digit year comparison code                     |  #
 ' #  |    - other changes:                                                  |  #
 ' #  |        = Input form significantly modified:                          |  #
 ' #  |            + CC date field is now a dynamic select box displaying    |  #
 ' #  |              every month and year combination for the next five      |  #
 ' #  |              years from today's date.                                |  #
 ' #  |            + CC number field only accepts a maximum of 20 characters |  #
 ' #  |            + added some simple client-side javascript to verify that |  #
 ' #  |              the form is completely filled out.                      |  #
 ' #  +----------------------------------------------------------------------+  #
 ' #                                                                            #
 ' #  +-------------+                                                           #
 ' #  | VERSION 3.0 |                                                           #
 ' #  |             +--------------------------------------------------------+  #
 ' #  |   released on: 9/20/2000                                             |  #
 ' #  |    - "VerifyCreditCardNumber" function changed to a VBScript class   |  #
 ' #  |      to replace spaghetti code with procedure blocks encapsulated    |  #
 ' #  |      into a class for neatness.                                      |  #
 ' #  |    - Numerous new features added:                                    |  #
 ' #  |        = regular expressions now clean and verify credit card        |  #
 ' #  |          numbers, eliminating the need for complex VBScript CCNumber |  #
 ' #  |          checks and length verifications.                            |  #
 ' #  |        = credit card and encoding/decoding constants introduced.     |  #
 ' #  |        = credit card information can be base64 encoded/decoded if    |  #
 ' #  |          desired for added security.                                 |  #
 ' #  |        = verification result outputs a number that can optionally be |  #
 ' #  |          decoded into a human readable message if necessary.         |  #
 ' #  |        = verification form encapsulated into the class as a method.  |  #
 ' #  |        = client-side form validation updated.                        |  #
 ' #  |        = Better date processing than before:                         |  #
 ' #  |            + No longer uses arrays to decypher date input            |  #
 ' #  |            + 4 digit years are forced by using the Year function     |  #
 ' #  |            + y2k/4 digit year compatible                             |  #
 ' #  +----------------------------------------------------------------------+  #
 ' #                                                                            #
 ' #                                                                            #
 ' # CONSTANTS:                                                                 #
 ' #  The following constants are aware to the class and have page scope,       #
 ' #  meaning you can use them in place of their corresponding values when      #
 ' #  using this class.                                                         #
 ' #                                                                            #
 ' #       +-------------------+--------------------+-------------------+       #
 ' #       |     CONSTANT      |    CONSTANT TYPE   |   ACTUAL VALUE    |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   MasterCard      |  credit card type  |         1         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   Visa            |  credit card type  |         2         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   AmericanExpress |  credit card type  |         3         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   DinersClub      |  credit card type  |         4         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   CarteBlanche    |  credit card type  |         4         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   Discover        |  credit card type  |         5         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   EnRoute         |  credit card type  |         6         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   JCB             |  credit card type  |         7         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   Base64Enc       |  encoding/decoding |         1         |       #
 ' #       |-------------------+--------------------+-------------------|       #
 ' #       |   Base64Dec       |  encoding/decoding |         2         |       #
 ' #       +-------------------+--------------------+-------------------+       #
 ' #                                                                            #
 ' #                                                                            #
 ' # PROPERTIES:                                                                #
 ' #  There are three properties, all of which are mandatory and must be        #
 ' #  set before calling any of the credit card processing methods              #
 ' #  (the EnDecryptCard() and ShowVerificationForm() methods are the only      #
 ' #  exceptions to this rule.)                                                 #
 ' #                                                                            #
 ' #  +---------------------+------------------------------------------------+  #
 ' #  |      PROPERTY       |              ACCEPTABLE INPUT                  |  #
 ' #  |---------------------+------------------------------------------------|  #
 ' #  |---------------------+------------------------------------------------|  #
 ' #  |  CardType           | Expects an integer or Constant that represents |  #
 ' #  |                     |   an acceptable credit card type.              |  #
 ' #  |---------------------+------------------------------------------------|  #
 ' #  |  CardNumber         | Expects a string that represents a credit card |  #
 ' #  |                     |   number to be verified before sending out for |  #
 ' #  |                     |   debit processing.                            |  #
 ' #  |---------------------+------------------------------------------------|  #
 ' #  |  CardExpirationDate | Expects a date that represents the expiration  |  #
 ' #  |                     |   date of the card.                            |  #
 ' #  +---------------------+------------------------------------------------+  #
 ' #                                                                            #
 ' #                                                                            #
 ' # METHODS:                                                                   #
 ' #  There are four methods exposed by this class.                             #
 ' #                                                                            #
 ' # +----------------------+-----------------------------+-------------------+ #
 ' # |        METHOD        |          ARGUMENTS          |      RETURNS      | #
 ' # |----------------------+-----------------------------+-------------------| #
 ' # |----------------------+-----------------------------+-------------------| #
 ' # | VerifyCard           | no arguments                | integer -         | #
 ' # |   - this method      |                             | returns 0 if card | #
 ' # |     verifies the     |                             | is valid or a     | #
 ' # |     card information |                             | number greater    | #
 ' # |     entered as       |                             | than zero if the  | #
 ' # |     class properties |                             | card failed       | #
 ' # |                      |                             | verification.     | #
 ' # |----------------------+-----------------------------+-------------------| #
 ' # | TranslateCardResults | 1.) integer - expects the   | string -          | #
 ' # |   - this method      |     return value of the     | translates the    | #
 ' # |     translates the   |     VerifyCard method       | return integer    | #
 ' # |     output of the    |                             | of the VerifyCard | #
 ' # |     VerifyCard       |                             | method and will   | #
 ' # |     method into a    |                             | output a human-   | #
 ' # |     message that is  |                             | readable message  | #
 ' # |     meaningful to a  |                             |                   | #
 ' # |     user             |                             |                   | #
 ' # |----------------------+-----------------------------+-------------------| #
 ' # | EnDecryptCard        | 1.) integer or constant -   | string -          | #
 ' # |   - this method      |     representing the base64 | returns a base64  | #
 ' # |     encrypts or      |     encoding or decoding    | encoded credit    | #
 ' # |     decrypts a       |     action to be performed  | card number or a  | #
 ' # |     credit card      |     by the method.          | decoded card      | #
 ' # |     number using the | 2.) string - expects a      | number.           | #
 ' # |     Base64 method    |     credit card number to   |                   | #
 ' # |                      |     encode or decode.       |                   | #
 ' # |----------------------+-----------------------------+-------------------| #
 ' # | ShowVerificationForm | no arguments                | string -          | #
 ' # |   - this method      |                             | returns the       | #
 ' # |     outputs the form |                             | input form        | #
 ' # |     that allows a    |                             |                   | #
 ' # |     user to enter    |                             |                   | #
 ' # |     credit card info |                             |                   | #
 ' # |     for validating   |                             |                   | #
 ' # +----------------------+-----------------------------+-------------------+ #
 ' #                                                                            #
 ' #                                                                            #
 ' # PROCEDURE CREDITS:                                                         #
 ' #  +-----------------------------+                                           #
 ' #  |  Tarun Bhushan              +----------------------------------------+  #
 ' #  |  tarunbhushan@hotmail.com                                            |  #
 ' #  |                                                                      |  #
 ' #  |    1.) TestLuhnCheckDigitAlgorithm                                   |  #
 ' #  |    2.) TranslateCardResults                                          |  #
 ' #  +----------------------------------------------------------------------+  #
 ' #                                                                            #
 ' #  +-----------------------------+                                           #
 ' #  |  webmaster@q-tec.org        +----------------------------------------+  #
 ' #  |                                                                      |  #
 ' #  |    1.) base64_encode                                                 |  #
 ' #  |    2.) base64_decode                                                 |  #
 ' #  |    3.) mimeencode                                                    |  #
 ' #  |    4.) mimedecode                                                    |  #
 ' #  +----------------------------------------------------------------------+  #
 ' #                                                                            #
 ' #                                                                            #
 ' # WARNINGS:                                                                  #
 ' #   VERIFICATION:                                                            #
 ' #    This class is meant to be a first step for validating credit cards. A   #
 ' #    credit card is not verified until it has been sent to a debit service   #
 ' #    or bank along with the amount of the charge. This class simply weeds    #
 ' #    out incorrect or fradulent card numbers that don't pass the given tests.#  
 ' #    Any card that is verified by using this class is good enough to send    #
 ' #    out for processing. Any that fail the verification by this class are    #
 ' #    not valid credit card numbers and don't need to be sent to a debit      #
 ' #    service.                                                                #
 ' #                                                                            #
 ' #   SECURITY:                                                                #
 ' #    This class provides Base64/MIME encoding of credit card numbers. You    #
 ' #    should realize that Base64 is not a secure encryption, it is simply     #
 ' #    a uniform way of representing a credit card number on any software      #
 ' #    platform or language and a guarantee that the card number can't be      #
 ' #    casually viewed without a Base64 Decoder. Base64 was created for use    #
 ' #    in email to allow binary strings to be converted to text-strings,       #
 ' #    allowing items like pictures to be safely emailed or posted on a        #
 ' #    newsgroup. Base64 encoding is provided for those who have no other      #
 ' #    security measures available when verifying entered credit card numbers  #
 ' #                                                                            #
 ' #    You should verify credit card numbers on a secure server to minimize    #
 ' #    the chance of interception by persons unknown. Also, most user's will   #
 ' #    not enter a card number if they don't see the little lock in the corner #
 ' #    of the browser, indicating that the site is secure. This is also        #
 ' #    verified by seeing HTTPS:// before a URL path.                          #
 ' #                                                                            #
 ' #    Remember that it's just bad business to email someone's credit card     #
 ' #    to another source for processing as it can be easily intercepted by     #
 ' #    third parties along the way. The card number may appear in server       #
 ' #    caches and logs as well, especially if the mail is accidently           #
 ' #    misrouted. You should also realize that most email and newsgroup        #
 ' #    viewers have Base64 Decoding built in to their architecture.            #
 ' #                                                                            #
 ' #    You should use better encryption methods to keep credit card data       #
 ' #    secure and you should send credit card data out for processing using    #
 ' #    the means given by your debit service or bank as these are usually      #
 ' #    more secure methods of transferring sensitive information.              #
 ' #                                                                            #
 ' #   DISCLAIMER:                                                              #
 ' #    I, nor any of the people involved in the creation of this class, will   #
 ' #    be liable for anything that happens as a result of using this class.    #
 ' #    You use this software at your own risk and by viewing, downloading,     #
 ' #    using, modifying or thinking about this class, you have agreed to the   #
 ' #    conditions above.                                                       #
 ' #                                                                            #
 ' #                                                                            #
 ' # KNOWN BUGS:                                                                #
 ' #  There are currently no known bugs associated with this class.             #
 ' #                                                                            #
 ' #  You can report bugs, alterations, additions, changes, questions,          #
 ' #  concerns or whatever about this class to the feedback area at the         #
 ' #  ASP Emporium:                                                             #
 ' #    http://www.aspEmporium.com/aspEmporium/feedback/index.asp               #
 ' #  Or to the developer's forum:                                              #
 ' #    http://www.aspEmporium.com/aspEmporium/forum/display_forum.asp?fid=7    #
 ' #                                                                            #
 ' #                                                                            #
 ' # SOFTWARE UPDATES AND MORE INFORMATION:                                     #
 ' #  This credit card verification system is currently in release version 3.0  #
 ' #  All of the previous versions of this software should be discarded in      #
 ' #  favor of this new system. None of the earlier versions of this software   #
 ' #  are supported. You are urged to keep your apps current. The most current  #
 ' #  version of this software can always be found at the ASP Emporium:         #
 ' #    http://www.aspEmporium.com/exampleView.asp?eg=CCVerification%5Fclass    #
 ' #                                                                            #
 ' #  This class requires no other files to work correctly.                     #
 ' #  You must be running VBScript 5.1 or better to use class files. It is      #
 ' #  recommended that you update to the latest version of the VBScript         #
 ' #  Scripting Engine, which at the time of this writing is 5.5, before using  #
 ' #  this class.                                                               #
 ' ##############################################################################


 ' #############################################
 ' #   class encryption/decryption constants   #
 ' #############################################
Const Base64Enc = 1, Base64Dec = 2


 ' #############################################
 ' #   class credit card type constants        #
 ' #############################################
Const MasterCard = 1, Visa = 2, AmericanExpress = 3, DinersClub = 4
Const CarteBlanche = 4, Discover = 5, EnRoute = 6, JCB = 7


Class CCVerification
	Private strCrdNum, IsClean, intCardType, dExpDate, Base64Chars

	 ' #################################################
	 ' #  Class Events                                 #
	 ' #    These routines are called when the         #
	 ' #    event occurs within the scripting engine.  #
	 ' #################################################

	Private Sub Class_Initialize()
		IsClean = false
		Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
				"abcdefghijklmnopqrstuvwxyz" & _
				"0123456789" & _
				"+/"
	End Sub

	Private Sub Class_Terminate()
		 Response.Write( vbCrLf & "<" & _
			"!-- CCVerification Class" & _
			" - http://www.aspEmporiu" & _
			"m.com/ --" & ">" & vbCrLf )
	End Sub

	 ' #################################################
	 ' #  Class Properties                             #
	 ' #    Properties are set by a developer in code  #
	 ' #    to configure and set up the class.         #
	 ' #################################################

	Public Property LET CardType(byVal intCardTp)
		dim re, bTmp
		set re = new RegExp
		with re
			.pattern = "^[1-7]$" : .global = true
			bTmp = .test(intCardTp)
		end with
		set re = nothing
		if bTmp then
			intCardType = CInt(intCardTp)
		else
			 ' set to MasterCard
			intCardType = 1
		end if
	End Property

	Public Property LET CardExpirationDate(byVal dCardExpDate)
		dExpDate = cDate(dCardExpDate)
	End Property

	Public Property LET CardNumber(byVal lngCardNumber)
		strCrdNum = CStr(lngCardNumber)
	End Property

	Public Property GET CardType()
		CardType = Int(intCardType)
	End Property

	Public Property GET CardExpirationDate()
		CardExpirationDate = cDate(dExpDate)
	End Property

	Public Property GET CardNumber()
		if NOT IsClean then strCrdNum = CleanCardNumber(CStr(strCrdNum))
		CardNumber = strCrdNum
	End Property

	 ' #################################################
	 ' #  Class Methods                                #
	 ' #    Methods are called by a developer to       #
	 ' #    perform an action                          #
	 ' #################################################

	Public Function VerifyCard()
		dim bPatternPassed, bLuhnCheckDigitPassed, bDatePassed
		dim strTmp, strTmp2, strTmp3
		strTmp  = CardNumber
		strTmp2 = CardExpirationDate
		strTmp3 = CardType
		If _
		    (Len(Trim(strTmp))  = 0) Or _
		    (Len(Trim(strTmp3)) = 0) Or _
		    (Len(Trim(strTmp2)) = 0) _
		Then
			VerifyCard = 1
			Exit Function
		End If

		bDatePassed = TestExpirationDate(strTmp2)
		If Not bDatePassed then VerifyCard = 8 : Exit Function

		bPatternPassed = TestCardPatterns(strTmp)
		If Not bPatternPassed then VerifyCard = 2 : Exit Function

		If strTmp3 <> EnRoute then
			bLuhnCheckDigitPassed = TestLuhnCheckDigitAlgorithm(strTmp)
			If Not bLuhnCheckDigitPassed then VerifyCard = 6 : Exit Function
		End If

		VerifyCard = 0
	End Function

	Public Function EnDecryptCard(byVal EncDecConstant, byVal cdNum)
		dim tmp
		tmp = cdNum
		if EncDecConstant = Base64Enc then
			 ' encode card number
			EnDecryptCard = base64_encode(tmp)
		elseif EncDecConstant = Base64Dec then
			 ' Decode card number
			EnDecryptCard = base64_decode(tmp)
		else
			 ' otherwise, return the card number
			EnDecryptCard = tmp
		end if
	End Function

	Public Function TranslateCardResults(byVal lngNumber)
		dim tmp
		Select Case cLng(lngNumber)
			Case 0    : tmp = "Card Validated Successfully!"
			Case 1    : tmp = "Missing Required Input! " & _
					  "Please fill in the entire form!"
			Case 2, 6 : tmp = "Invalid Card Number!"
			Case 8    : tmp = "Bad Expiration Date!"
			Case Else : tmp = "An Unknown Error Has Occurred!"
		End Select
		TranslateCardResults = tmp
	End Function

	Public Function ShowVerificationForm()
		dim tmp, m1, y1
		tmp = ""
		tmp = tmp & "<CENTER>" & vbCrLf
		tmp = tmp & "<TABLE ALIGN=CENTER BGCOLOR=""#60786B"" CELLPADDING=1 "
		tmp = tmp & "CELLSPACING=1 BORDER=0>" & vbCrLf & vbTab
		tmp = tmp & "<TR>" & vbCrLf & vbTab & vbTab & "<TD "
		tmp = tmp & "BGCOLOR=""#FFFFEE"">" & vbCrLf
		tmp = tmp & vbTab & vbTab & "<form action="""
		tmp = tmp & Request.ServerVariables("SCRIPT_NAME")
		tmp = tmp & """ method=""post"" onsubmit=""javascript:"
		tmp = tmp & "return checkCCInput(this)"">"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab
		tmp = tmp & "<table>" & vbCrLf & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<tr>" & vbCrLf & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & vbTab & "<td>Credit Card Type: </td>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<td>" & vbCrLf & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & vbTab
		tmp = tmp & "<select name=""cc_type"" size=1>" & vbCrLf
		tmp = tmp & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""1"">Mastercard</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""2"">Visa</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""3"">American Express</option>" 
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""4"">Diner's Club</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""4"">Carte Blanche</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""5"">Discover</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""6"">enRoute</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<option value=""7"">JCB</option>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "</select>" & vbCrLf & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & "</td>" & vbCrLf & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & "</tr>" & vbCrLf & vbTab & vbTab
		tmp = tmp & vbTab & vbTab
		tmp = tmp & "<tr>" & vbCrLf & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<td>Credit Card Number:</td>" & vbCrLf
		tmp = tmp & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "<td>" & vbCrLf & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<input type=text name=""cc_num"" "
		tmp = tmp & "value="""" size=20 maxlength=20>" & vbCrLf & vbTab
		tmp = tmp & vbTab & vbTab & vbTab & vbTab & "</td>" & vbCrLf
		tmp = tmp & vbTab & vbTab & vbTab & vbTab & "</tr>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & "<tr>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "<td>Expiration Date: mm/yyyy</td>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab 
		tmp = tmp & "<td>" & vbCrLf & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & "<SELECT NAME=""cc_exp"" SIZE=1>" & vbCrLf
		for y1 = 0 to 5 step 1
			for m1 = 1 to 12 step 1
				if m1 < month(date()) and y1 = 0 then
				else
					if len(m1) = 1 then m1 = "0" & cstr(m1)
					tmp = tmp & vbTab & vbTab & vbTab & vbTab
					tmp = tmp & vbTab & vbTab
					tmp = tmp & "<OPTION VALUE=""" & m1 & "/"
					tmp = tmp & year(date()) + y1 & """>"
					tmp = tmp & m1 & "/" & year(date()) + y1
					tmp = tmp & "</OPTION>" & vbCrLf
				end if
			next
		next
		tmp = tmp & vbTab & vbTab & vbTab & vbTab & vbTab & "</SELECT>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "</td>" & vbCrLf & vbTab & vbTab & vbTab & vbTab & "</tr>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & "<tr>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "<td colspan=2 align=center>"
		tmp = tmp & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab
		tmp = tmp & "<input type=submit value=""perform "
		tmp = tmp & "verification"">" & vbCrLf & vbTab & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & "</td>" & vbCrLf & vbTab & vbTab
		tmp = tmp & vbTab & vbTab & "</tr>" & vbCrLf
		tmp = tmp & vbTab & vbTab & vbTab & "</table>" & vbCrLf & vbTab
		tmp = tmp & vbTab & "</form>" & vbCrLf & vbTab & vbTab & "</TD>"
		tmp = tmp & vbCrLf & vbTab & "</TR>" & vbCrLf & "</TABLE>"
		tmp = tmp & vbCrLf & "</CENTER>" & vbCrLf
		ShowVerificationForm = tmp
	End Function

	 ' #################################################
	 ' #  Class Routines                               #
	 ' #    These are internal routines that are used  #
	 ' #    by the class to validate card information  #
	 ' #################################################

	Private Function CleanCardNumber(byVal ccNum)
		 ' remove white space, dashes (-) and non digits
		dim re : set re = new RegExp
		ccNum = cStr(ccNum)
		with re
			.global = true : .ignorecase = false : .multiline = false
			.pattern = "(\-|\s|\D)*"
			CleanCardNumber = .replace(ccNum, "")
		end with
		set re = nothing
		IsClean = true
	End Function

	Private Function TestCardPatterns(byVal lngCardNumber)
		 ' Introducing regular expressions to verify card numbers
		 ' shortened the amount of coding needed to properly
		 ' test cards. This pattern eliminated approximately
		 ' 60 lines of VBScript that used to test card prefixes
		 ' and lengths.

		dim re : set re = new RegExp : lngCardNumber = CStr(lngCardNumber)
		dim strPattern : strPattern = ""

		 ' this regular expression pattern tests every
		 ' possible credit card type's prefix and length 
		 ' combination and rejects any that don't fit the 
		 ' criteria. Each test for a particular card is 
		 ' on a separate line below:
		
		Select Case CardType
			Case MasterCard
				strPattern = strPattern & _
				  "^(?:(?:[5][1-5])(?:\d{14}))$"
			Case Visa
				strPattern = strPattern & _
				  "^(?:(?:[4])(?:\d{12}|\d{15}))$"
			Case AmericanExpress
				strPattern = strPattern & _
				  "^(?:(?:[3][4|7])(?:\d{13}))$"
			Case CarteBlanche, DinersClub
				strPattern = strPattern & _
				  "^(?:(?:[3](?:[0][0-5]|[6|8]))(?:\d{11,12}))$"
			Case Discover
				strPattern = strPattern & _
				  "^(?:(?:6011)(?:\d{12}))$"
			Case EnRoute
				strPattern = strPattern & _
				  "^(?:(?:[2](?:014|149))(?:\d{11}))$"
			Case JCB
				strPattern = strPattern & _
				  "^(?:(?:(?:2131|1800)(?:\d{11}))$|^(?:(?:3)(?:\d{15})))$"
			Case Else
				TestCardPatterns = false
				exit function
		End Select
		with re
			.global = true     : .ignorecase = false 
			.multiline = false : .pattern = strPattern
			TestCardPatterns = .test(lngCardNumber)
		end with
		set re = nothing
	End Function

	Private Function TestLuhnCheckDigitAlgorithm(byVal strCCNumber)

		 ' #################################################################
		 ' # This function was provided by: Tarun Bhushan                  #
		 ' #                                tarunbhushan@hotmail.com       #
		 ' #                                                               #
		 ' # Check Digit Algorithm.                                        #
		 ' # This is the algorithm step by step:                           #
		 ' #                                                               #
		 ' # 1. Reverse the credit card number and split into              #
		 ' #    individual digits.  The reason for this is that            #
		 ' #    for an odd digit count, the digits of interest             #
		 ' #    are the even ones.  For an even digit count,               #
		 ' #    the digits of interest are the odd ones.                   #
		 ' #    Reversing a number with an odd number of digits            #
		 ' #    leaves the same digits in the even positions.              #
		 ' #    Reversing a number with an even number of digits           #
		 ' #    swaps the even and odd positions.                          #
		 ' #                                                               #
		 ' #    For example:                                               #
		 ' #     (Odd digit count)             (Odd digit count reversed)  #
		 ' #     123456789                     987654321                   #
		 ' #     OEOEOEOEO                     OEOEOEOEO                   #
		 ' #      \      \_____________________/      /                    #
		 ' #       \_________________________________/                     #
		 ' #                       |                                       #
		 ' # Notice that both 1 and 8 are in an odd and even position,     #
		 ' # respectively, in both cases.                                  #
		 ' #                                                               #
		 ' #     (Even digit count)            (Even digit count reversed) #
		 ' #     12345678                      87654321                    #
		 ' #     OEOEOEOE                      OEOEOEOE                    #
		 ' #     --------                      --------                    #
		 ' #      \      \____________________/      /                     #
		 ' #       \________________________________/                      #
		 ' #                       |                                       #
		 ' #     Notice that 1 changes from an odd position to an even     #
		 ' #     position, and 8 changes from an even position to an odd   #
		 ' #     position.                                                 #
		 ' #                                                               #
		 ' #                                                               #
		 ' # 2. Multiply every even numbered digit by 2.  If the result    #
		 ' #    is greater than 9, 9 is subtracted from the product.       #
		 ' #    The original even digit is replaced with the result of     #
		 ' #    these calculations.                                        #
		 ' #                                                               #
		 ' #                                                               #
		 ' # 3. Add all the digits together (the original odd digits and   #
		 ' #    the replaced even digits)                                  #
		 ' #                                                               #
		 ' #                                                               #
		 ' # 4. Divide this sum by 10.  If the remainder is zero, the      #
		 ' #    number is valid. Otherwise, the number is invalid.         #
		 ' #                                                               #
		 ' #################################################################

		dim dCCNLen : dCCNLen = Int(Len(strCCNumber))
		dim strRevCCNumber : strRevCCNumber = strReverse(strCCNumber)
		dim dDigitTimesTwo, i
		dim dValNumber : dValNumber = Int(0)
		dim dCCNArray() : ReDim dCCNArray(dCCNLen)

		' Step 1
		For i = 1 to dCCNLen
			dCCNArray(i) = Int(Mid(strRevCCNumber,i,1))
		Next

		' Step 2
		For i = 2 to dCCNLen Step 2
			dDigitTimesTwo = Int(dCCNArray(i) * 2)
			if dDigitTimesTwo > 9 Then
				dCCNArray(i) = Int(dDigitTimesTwo - 9)
			else
				dCCNArray(i) = Int(dDigitTimesTwo)
			end if
		Next

		' Step 3
		For i = 1 to dCCNLen
			dValNumber = dValNumber + Int(dCCNArray(i))
		Next

		' Step 4
		dValNumber = dValNumber Mod 10

		' This is it
		if dValNumber <> 0 Then 
			TestLuhnCheckDigitAlgorithm = false : Exit Function
		end if
		TestLuhnCheckDigitAlgorithm = true
	End Function

	Private Function TestExpirationDate(byVal dExpiresOn)
		 ' Better date processing than before:
		 '  - No longer uses arrays to decypher date input
		 '  - 4 digit years are forced by using the Year function
		 '  - y2k compatible

		If NOT IsDate(dExpiresOn) Then
			TestExpirationDate = false : Exit Function
		End If
		dim dMonth, dYear, dThisYear, dThisMonth
		dMonth     = month(dExpiresOn)
		dThisMonth = month(date)
		dYear      = year(dExpiresOn)  ' outputs a 4 digit year
		dThisYear  = year(date)	       ' outputs a 4 digit year
		if dYear = dThisYear then
			 ' check month
			if dMonth >= dThisMonth then
				 ' valid but may be expiring 
				 ' at the end of the month
				TestExpirationDate = true  : Exit Function
			else
				 ' bad date or expired
				TestExpirationDate = false : Exit Function
			end if
		elseif dYear > dThisYear then
			 ' valid
			TestExpirationDate = true  : Exit Function
		else
			 ' bad date or expired
			TestExpirationDate = false : Exit Function
		end if
	End Function

	Public Function base64_encode(byVal strIn)
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 3
			c1 = Asc(Mid(strIn, n, 1))
			c2 = Asc(Mid(strIn, n + 1, 1) + Chr(0))
			c3 = Asc(Mid(strIn, n + 2, 1) + Chr(0))
			w1 = Int(c1 / 4) : w2 = (c1 And 3) * 16 + Int(c2 / 16)
			If Len(strIn) >= n + 1 Then 
				w3 = (c2 And 15) * 4 + Int(c3 / 64) 
			Else 
				w3 = -1
			End If
			If Len(strIn) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode(w1) + mimeencode(w2) + _
					  mimeencode(w3) + mimeencode(w4)
		Next
		base64_encode = strOut
	End Function

	Public Function base64_decode(byVal strIn)
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len(strIn) Step 4
			w1 = mimedecode(Mid(strIn, n, 1))
			w2 = mimedecode(Mid(strIn, n + 1, 1))
			w3 = mimedecode(Mid(strIn, n + 2, 1))
			w4 = mimedecode(Mid(strIn, n + 3, 1))
			If w2 >= 0 Then _
				strOut = strOut + _
					Chr(((w1 * 4 + Int(w2 / 16)) And 255))
			If w3 >= 0 Then _
				strOut = strOut + _
					Chr(((w2 * 16 + Int(w3 / 4)) And 255))
			If w4 >= 0 Then _
				strOut = strOut + _
					Chr(((w3 * 64 + w4) And 255))
		Next
		base64_decode = strOut
	End Function

	Private Function mimeencode(byVal intIn)
		If intIn >= 0 Then 
			mimeencode = Mid(Base64Chars, intIn + 1, 1) 
		Else 
			mimeencode = ""
		End If
	End Function

	Private Function mimedecode(byVal strIn)
		If Len(strIn) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr(Base64Chars, strIn) - 1
		End If
	End Function
End Class
%>