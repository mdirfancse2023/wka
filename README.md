 Public Function CheckForWKAFile(ByVal strInputDataFile As String, ByVal strBranchNo As String) As String
        Dim iFileNum As Integer
        Dim strCurrentLineFromDataFile As String
        Dim currLinePos As Integer
        Dim ProceedFlag As String = "NOERROR"
        Dim strPhoneno As String

        Dim strVar As String
        Dim strFlag As String
        Dim i As Int32
        Dim sLogFileName As String
        Dim arrayListOfErrors As New ArrayList()
        Dim strIDtypeCif As String
        'Start of IR 16040002
        Dim strSecIDtypeCif As String
        Dim trimmedCIFSecID As String
        Dim strSecIDnoCif As String
        Dim strSecIDNumberFullCif As String
        Dim strSecIDNumberFulltrimCif As String
        Dim strSecIDNumberfirstCharCif As String
        Dim strSecIDNumberlastCharCif As String
        'End of IR 16040002
        'start of IR 18050054
        Dim CKYC_CITIZENSHIP As String
        Dim CKYC_OCCUPANCY_TYPE As String
        Dim CKYC_ID_exp_DT As String
        Dim exp_DT1 As String
        Dim CKYC_ADDRESS_TYPE As String
        Dim CKYC_ADD_PROOF As String
        Dim CKYC_OTHER_ADD_PROOF As String
        Dim CKYC_TAX_LINE1 As String
        Dim CKYC_TAX_LINE2 As String
        Dim CKYC_TAX_LINE3 As String
        Dim CKYC_ANEX_DISTRICT As String
        Dim CKYC_ANEX_PIN As String
        Dim CKYC_ANEX_STATE As String
        Dim CKYC_TAX_COUNTRY As String
        Dim effective_DT As String
        Dim CKYC_RP1_PASS_EXP_DT As String
        Dim CKYC_RELIGION As String
        Dim CKYC_OTHER_RELIGION As String
        Dim CKYC_CATEGORY As String
        Dim CKYC_ORG_NAME As String
        Dim CKYC_DESIGNATION As String
        Dim CKYC_PREFIX2 As String
        Dim CKYC_FATHER_FNAME As String
        Dim CKYC_PREFIX3 As String
        Dim CKYC_MOTHER_FNAME As String



        Dim CKYC_PREFIX1 As String
        Dim CKYC_MAIDEN_FNAME As String
        Dim CKYC_MAIDEN_MNAME As String
        Dim CKYC_MAIDEN_LNAME As String



        Dim CKYC_FATHER_MNAME As String
        Dim CKYC_FATHER_LNAME As String


        Dim CKYC_MOTHER_MNAME As String
        Dim CKYC_MOTHER_LNAME As String



        Dim CKYC_CRES As String
        Dim CKYC_TIN As String
        Dim CKYC_BIRTH_PLACE As String

        Dim CKYC_BIRTH_COUNTRY As String

        Dim CKYC_TAX_CITY As String
        Dim CKYC_TAX_PIN As String
        Dim CKYC_TAX_STATE As String

        Dim CKYC_ANEX_LINE1 As String
        Dim CKYC_ANEX_LINE2 As String
        Dim CKYC_ANEX_LINE3 As String

        Dim CKYC_ANEX_CITY As String

        Dim CKYC_ANEX_COUNTRY As String
        Dim CKYC_INCOME As String
        Dim CKYC_NET_WORTH As String
        Dim CKYC_OCCUPATION As String
        Dim CKYC_NATURE_BUSINESS As String
        Dim CKYC_POLITICAL_FLAG As String

        Dim KHostDate As String
        Dim CKYC_KYC_NO1 As String
        Dim CKYC_PER1_TYPE As String
        Dim CKYC_PER1_PREFIX As String
        Dim CKYC_PER1_NAME1 As String
        Dim CKYC_PER1_MIDNAME As String
        Dim CKYC_PER1_NAME2 As String
        Dim strFatherPreFlag As String
        Dim strFatherPreFlag2 As String
        'end of IR 18050054
        'start of 18050127
        Dim AnnualIncome As String
        Dim Consent_code As String
        Dim Date_of_ConsentRevocation As String
        'end of 18050127
        'Start of IR 19030186
        Dim Risk_category As String
        'End of IR 19030186
        'Start of IR 18100014
        Dim Occup_Code As String
        'Start of IR 22020030
        'Dim Occup_Desc As String
        'End of IR 22020030
        Dim Emp_Pan As String
        'End of IR 18100014
        'Start of IR 19080031
        Dim CIFNo As String
        'End of IR 19080031
        'Start of IR 19020158
        Dim xn1 As XmlNode
        Dim xn2 As XmlNode
        Dim xn3 As XmlNode
        Dim xn4 As XmlNode
        Dim Str_TAX_PAN As String
        'End of IR 19020158
        Dim strIDNumberFullCif As String
        Dim strIDNumberFulltrimCif As String
        Dim strIDNumberfirstCharCif As String
        Dim strIDNumberlastCharCif As String
        Dim strIDnoCif As String
        Dim trimmedCIFID As String
        Dim strDrivingLic As String
        Dim strDrivingLic1 As Int32
        Dim strwelcomekitsweepflag As String
        Dim count As Int32 = 0

        'Start of IR 20090117
        Dim strForm60 As String
        Dim Form60_SubmissionDate As String
        Dim form60SubDate As Long = 0
        Dim Form60_TransactionDate As String
        Dim Form60_AgricultureIncome As String
        Dim Form60_Other_AgricultureIncome As String
        Dim Form60_PanApplid As String
        Dim Form60_AckPanApplied As String
        Dim Form60_PanAppliedDate As String
        'End of IR 20090117

        'Start of IR 22120029
        Dim source_fund As String
        'End of IR 22120029
        'Start of IR 23020024
        Dim idissuedt As String
        'End of IR 23020024
        'Start of IR 23020051
        Dim dobWKA As String
        'End of Ir 23020051
        'Start of IR 24030056
        Dim name_onCard As String
        'End of IR 24030056
        Try
            iFileNum = FreeFile()
            Try
                'START OF IR 20100041(Path Manipulation)
                iFileNum = validate.Canonicalize(iFileNum)
                strInputDataFile = validate.Canonicalize(strInputDataFile)
                'END OF IR 20100041 (Path Manipulation)
                FileOpen(iFileNum, strInputDataFile, OpenMode.Input)
            Catch ex As Exception
                FileOpen(iFileNum, strInputDataFile, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            End Try

            While Not EOF(iFileNum)
                LineNo = LineNo + 1
                
                Dim errorString = Array.CreateInstance(GetType(String), 102)
                

                Try
                 
                    currLinePos = 0
                    strCurrentLineFromDataFile = LineInput(iFileNum)
                   
                    totalWKAlen = strCurrentLineFromDataFile.Length
                   
                    If (strCurrentLineFromDataFile.Length() > 1046) Then
                        
                        errorString.SetValue("L", 0)
                        ProceedFlag = "N" 'Signifies invalid Record Length
                    End If


                    currLinePos = 0


                     'validate branch code
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 5)
                    If strVar = strBranchNo Then
                        strFlag = "Y"
                    Else
                        strFlag = "N"
                    End If
                    If strFlag = "N" Then
                        ProceedFlag = "N"
                        errorString.SetValue("B", 2)
                    End If
                    currLinePos = currLinePos + 5

                    'Validate WK Account No
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 17)
                    strFlag = ValidateNumber(strVar)
                    If strFlag = "N" Or strVar.Trim.Length <> 17 Then
                        ProceedFlag = "N"
                        errorString.SetValue("A", 3)
                    Else
                        If validateCheckDigit(strVar) = "N" Or strVar = "00000000000000000" Then
                            ProceedFlag = "N"
                            errorString.SetValue("A", 3)
                        End If
                    End If
                    currLinePos = currLinePos + 17

                    'validate title code
                    currLinePos = currLinePos + 2

                    'validate customers first name
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)
                    'Start of IR 20010091
                    firstNameWka = currLinePos
                    'End of IR 20010091
                    'Start of IR 23120129
                    'If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                    '    ProceedFlag = "N"
                    '    errorString.SetValue("F", 4)
                    'End If
                    If (ValidateAlphabetsQuotes(strVar) = "N") Or ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("F", 4)
                    End If
                    'End of IR 23120129
                    currLinePos = currLinePos + 40

                    'validate customers Middle name
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)

                    currLinePos = currLinePos + 40
                    'Start of IR 23120129

                    If (ValidateAlphabetsQuotes(strVar) = "N") Then
                        ProceedFlag = "N"
                        errorString.SetValue("Middle_Name", 99)
                    End If
                    'End of IR 23120129

                    'validate customers Last name
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)
                    'Start of IR 20010091
                    lastNameWka = currLinePos
                    
                    If (ValidateAlphabetsDots(strVar) = "N") Or ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("LN", 5)
                    End If
                    'End of IR 23120129
                    currLinePos = currLinePos + 40

                    'validate gender code


                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 1)
                    If (Not (strVar = "M" Or strVar = "F" Or strVar = "T")) Then
                        errorString.SetValue("G", 6) 'Signifies invalid   GenderCode
                        ProceedFlag = "N"
                    End If
                    currLinePos = currLinePos + 1


                    'validate Address line1
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)
                    If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("A1", 7)
                    End If
                    currLinePos = currLinePos + 40

                    ''validate Address line2
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)

                    currLinePos = currLinePos + 40

                    'validate Address line3
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)
                    If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("A3", 8)
                    End If
                    currLinePos = currLinePos + 40

                    ''validate Address line4
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 40)
                    If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("A4", 9)
                    End If
                    currLinePos = currLinePos + 40

                    'validate postal code
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 8)
                    strFlag = ValidateNumber(strVar)
                    If strFlag = "N" Then
                        ProceedFlag = "N"
                        errorString.SetValue("P", 10)
                    End If
                    currLinePos = currLinePos + 8

                    'validate city code
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 3)
                    If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("C", 11)
                    End If
                    currLinePos = currLinePos + 3

                    'validate state code
                    strVar = strCurrentLineFromDataFile.Substring(currLinePos, 2)
                    If ((strVar.Substring(0, 1) = " ") Or (strVar.Trim() = "")) Then
                        ProceedFlag = "N"
                        errorString.SetValue("S", 12)
                    End If
                    currLinePos = currLinePos + 2

                    'validate phone number
                    Try
                        strPhoneno = strCurrentLineFromDataFile.Substring(currLinePos, 12)

                        If (Not ((strPhoneno.Length = 12) And (strPhoneno.Trim = ""))) Then
                            If (strPhoneno.Substring(0, 1) = " ") Then
                                ProceedFlag = "N" 'Mobile number should not start with space
                                errorString.SetValue("ECA", 17)
                            Else

                                If (ValidateNumber(strPhoneno) = "N") Then
                                    ProceedFlag = "N"
                                    errorString.SetValue("PH", 13)

                                    Dim s As String = strPhoneno.Trim
                                ElseIf ((strPhoneno.TrimEnd.Length < 5)) Then
                                    ProceedFlag = "N" 'Signifies mobile number should not be less than 5 digits
                                    errorString.SetValue("ACA", 14)

                                ElseIf (strPhoneno.Substring(0, 1) = 0) Then
                                    ProceedFlag = "N" 'Signifies mobile number should not start with 0
                                    errorString.SetValue("BCA", 15)

                                ElseIf (1) Then
                                    Dim flag As Boolean = False
                                    i = 0

                                    While (i < (strPhoneno.Trim.Length) - 1)

                                        If (strPhoneno.Trim.Substring(i, 1) = strPhoneno.Trim.Substring(i + 1, 1)) Then

                                            flag = False
                                        Else
                                            flag = True
                                            Exit While
                                        End If
                                        i = i + 1
                                    End While
                                    If flag = True Then

                                    Else
                                        ProceedFlag = "N" 'Signifies all digits should not be same in a mobile number
                                        errorString.SetValue("DCA", 16)
                                    End If
                                End If
                            End If
                        End If
                    Catch e As Exception
                        ProceedFlag = "N" 'Mobile number should not start with space
                        errorString.SetValue("ECA", 17)
                    End Try
                    If strPhoneno.Trim = "" Then
                        ProceedFlag = "N"
                        errorString.SetValue("PH", 13)
                    End If
                    currLinePos = currLinePos + 12
                    'vadlidate id Type


                    strIDtypeCif = strCurrentLineFromDataFile.Substring(currLinePos, 2)

                    trimmedCIFID = strIDtypeCif.Trim()
                    If (trimmedCIFID <> "") Then
                        'Start of IR 16040002
                        'strFlag = ValidateNumber(strIDtypeCif)
                        strFlag = ValidateAlphaNumeric(strIDtypeCif)
                        'End of IR 16040002

                        If strFlag = "N" Then
                            ProceedFlag = "N"
                            errorString.SetValue("I", 18)
                        End If
                        'Start of IR 16040002
                    Else
                        ProceedFlag = "N"
                        errorString.SetValue("I", 18)
                        'End of IR 16040002
                    End If

                    currLinePos = currLinePos + 2

                    'validate id number
                    strIDnoCif = strCurrentLineFromDataFile.Substring(currLinePos, 15)
                    'Start of IR 23020024
                    'If trimmedCIFID <> "10" And trimmedCIFID <> "02" Then
                    If trimmedCIFID <> "44" And trimmedCIFID <> "10" And trimmedCIFID <> "02" Then
                        'End of IR 23020024
                        If ((strIDnoCif.Substring(0, 1) = " ") Or (strIDnoCif.Trim() = "")) Then
                            ProceedFlag = "N"
                            errorString.SetValue("EK", 21)
                        End If

                    End If

