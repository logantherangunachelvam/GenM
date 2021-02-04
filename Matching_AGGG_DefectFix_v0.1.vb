'----------------------------------------------
' Import table code for PBB 
' Prepared by: Logantheran K Gunachelvam
' Last edited: 4th February 2021
'----------------------------------------------

If importTableType = "MR_PBB_PBBPYT" Then

        varfile = importTablePath
            propertyType = RegexMatches(fso.GetBaseName(importTablePath), "(?!\w+_)[a-zA-Z0-9]+")
            merchantReportDate = RegexMatches(fso.GetBaseName(importTablePath), "[0-9]{8}(?=_MR)")
    
            'Initialize Excel application to load variables for IF and FOR statements
            Set xlApp = New Excel.Application
            xlApp.Visible = False
    
            Set xlWb = xlApp.Workbooks.Open(varfile)
            Set xlWs = xlWb.Worksheets(1)    
            
            columnValueCheck = xlWs.Cells(1, 14).Value
            miscColumnStart = 13
            miscColumnEnd = xlWs.Cells(1, 1).End(xlToRight).Column
            miscColumnCounter = 1
            miscColumnStop = 16
            
            xlWb.Close False
            xlApp.Quit
    
            Set xlApp = Nothing
            Set xlWb = Nothing
            Set xlWs = Nothing
    
         If fileCounter = 0 Then
        
            'Validate if source file contains less than 16 columns
            'If less than 16 columns, sqlQuery is built dynamically by appending the sqlQuery text
            If miscColumnEnd < 16 Then
            
                'Build the initial fixed columns in the query string
                sqlQuery = "SELECT [" & importTableSheet & "$].F1 AS MerchantID, [" & importTableSheet & "$].F2 AS RefNum, [" & importTableSheet & "$].F3 AS InvoiceNum, [" & importTableSheet & "$].F4 AS CardNum, [" & importTableSheet & "$].F5 AS Transaction_Date, [" & importTableSheet & "$].F6 AS Transaction_Time, [" & importTableSheet & "$].F7 AS Transaction_Code, [" & importTableSheet & "$].F8 AS Transaction_Type, " & _
                                "[" & importTableSheet & "$].F9 AS Number_Submitted, [" & importTableSheet & "$].F10 AS Amount, [" & importTableSheet & "$].F11 AS RowNum, [" & importTableSheet & "$].F12 AS NetAmount, "
                                
                'Loop to add the column where it EXISTS in source file
                For i = miscColumnStart To miscColumnEnd
                    sqlQuery = sqlQuery + "[" & importTableSheet & "$].F" & i & " AS Misc_Col" & miscColumnCounter & ", "
                    miscColumnCounter = miscColumnCounter + 1
                    miscColumnEnd = miscColumnEnd + 1
                Next i
                
                'Loop to add DUMMY columns that DOES NOT EXIST in source file                   
                For i = miscColumnEnd To miscColumnStop
                    sqlQuery = sqlQuery + "NULL AS Misc_Col" & miscColumnCounter & ", "
                    miscColumnCounter = miscColumnCounter + 1
                Next i
            
                'Build the final piece of the query string by indicating the DAO Excel object to import, and execute the query
                sqlQuery = sqlQuery + "'" & merchantReportDate & "' AS Merchant_Report_Date " & _
                        "INTO " & importTable & " " & _
                        "FROM [Excel 12.0 Xml;HDR=No;Database=" & importTablePath & "].[" & importTableSheet & "$] " & _
                        "WHERE (([" & importTableSheet & "$].F1)<>'Merchant Account ID');"
                    
                CurrentDb.Execute sqlQuery, dbFailOnError

            Else    'If it's more or equal to 16 columns, we only want to pick the first 16 columns and execute the query
            
                sqlQuery = "SELECT [" & importTableSheet & "$].F1 AS MerchantID, [" & importTableSheet & "$].F2 AS RefNum, [" & importTableSheet & "$].F3 AS InvoiceNum, [" & importTableSheet & "$].F4 AS CardNum, [" & importTableSheet & "$].F5 AS Transaction_Date, [" & importTableSheet & "$].F6 AS Transaction_Time, [" & importTableSheet & "$].F7 AS Transaction_Code, [" & importTableSheet & "$].F8 AS Transaction_Type, " & _
                                "[" & importTableSheet & "$].F9 AS Number_Submitted, [" & importTableSheet & "$].F10 AS Amount, [" & importTableSheet & "$].F11 AS RowNum, [" & importTableSheet & "$].F12 AS NetAmount, [" & importTableSheet & "$].F13 AS Misc_Col1, [" & importTableSheet & "$].F14 AS Misc_Col2, [" & importTableSheet & "$].F15 AS Misc_Col3, [" & importTableSheet & "$].F16 AS Misc_Col4, '" & merchantReportDate & "' AS Merchant_Report_Date " & _
                                "INTO " & importTable & " " & _
                                "FROM [Excel 12.0 Xml;HDR=No;Database=" & importTablePath & "].[" & importTableSheet & "$] " & _
                                "WHERE (([" & importTableSheet & "$].F1)<>'Merchant Account ID');"
                    
                CurrentDb.Execute sqlQuery, dbFailOnError
                    
            End If

        'This ELSE loop is triggered
        Else

            'Validate if source file contains less than 16 columns
            'If less than 16 columns, sqlQuery is built dynamically by appending the sqlQuery text
            If miscColumnEnd < 16 Then

                'Build the initial fixed columns in the query string
                sqlQuery = "INSERT INTO " & importTable & " SELECT [" & importTableSheet & "$].F1 AS MerchantID, [" & importTableSheet & "$].F2 AS RefNum, [" & importTableSheet & "$].F3 AS InvoiceNum, [" & importTableSheet & "$].F4 AS CardNum, [" & importTableSheet & "$].F5 AS Transaction_Date, [" & importTableSheet & "$].F6 AS Transaction_Time, [" & importTableSheet & "$].F7 AS Transaction_Code, [" & importTableSheet & "$].F8 AS Transaction_Type, " & _
                                "[" & importTableSheet & "$].F9 AS Number_Submitted, [" & importTableSheet & "$].F10 AS Amount, [" & importTableSheet & "$].F11 AS RowNum, [" & importTableSheet & "$].F12 AS NetAmount, "

                'Loop to add the column where it EXISTS in source file
                For i = miscColumnStart To miscColumnEnd
                    sqlQuery = sqlQuery + "[" & importTableSheet & "$].F" & i & " AS Misc_Col" & miscColumnCounter & ", "
                    miscColumnCounter = miscColumnCounter + 1
                    miscColumnEnd = miscColumnEnd + 1
                Next i

                'Loop to add DUMMY columns that DOES NOT EXIST in source file                   
                For i = miscColumnEnd To miscColumnStop
                    sqlQuery = sqlQuery + "NULL AS Misc_Col" & miscColumnCounter & ", "
                    miscColumnCounter = miscColumnCounter + 1
                Next i

                 'Build the final piece of the query string by indicating the DAO Excel object to import, and execute the query
                sqlQuery = sqlQuery + "'" & merchantReportDate & "' AS Merchant_Report_Date " & _
                        "FROM [Excel 12.0 Xml;HDR=No;Database=" & importTablePath & "].[" & importTableSheet & "$] " & _
                        "WHERE (([" & importTableSheet & "$].F1)<>'Merchant Account ID');"
                    
                CurrentDb.Execute sqlQuery, dbFailOnError

            Else 
            
                'Import data from source file up to 16 columns	
                sqlQuery = "INSERT INTO " & importTable & " SELECT [" & importTableSheet & "$].F1 AS MerchantID, [" & importTableSheet & "$].F2 AS RefNum, [" & importTableSheet & "$].F3 AS InvoiceNum, [" & importTableSheet & "$].F4 AS CardNum, [" & importTableSheet & "$].F5 AS Transaction_Date, [" & importTableSheet & "$].F6 AS Transaction_Time, [" & importTableSheet & "$].F7 AS Transaction_Code, [" & importTableSheet & "$].F8 AS Transaction_Type, " & _
                                "[" & importTableSheet & "$].F9 AS Number_Submitted, [" & importTableSheet & "$].F10 AS Amount, [" & importTableSheet & "$].F11 AS RowNum, [" & importTableSheet & "$].F12 AS NetAmount, [" & importTableSheet & "$].F13 AS Misc_Col1, [" & importTableSheet & "$].F14 AS Misc_Col2, [" & importTableSheet & "$].F15 AS Misc_Col3, [" & importTableSheet & "$].F16 AS Misc_Col4, '" & merchantReportDate & "' AS Merchant_Report_Date " & _
                                "FROM [Excel 12.0 Xml;HDR=No;Database=" & importTablePath & "].[" & importTableSheet & "$] " & _
                                "WHERE (([" & importTableSheet & "$].F1)<>'Merchant Account ID');"
                   
                CurrentDb.Execute sqlQuery, dbFailOnError
            
        End If
    
    End If
