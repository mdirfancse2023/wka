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
