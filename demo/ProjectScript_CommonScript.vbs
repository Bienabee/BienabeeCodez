Dim STPass, STFail, STNA, STNC, STNR

Function ActionCanExecute(ActionName)
  'Use ActiveModule and ActiveDialogName to get
  'the current context.
On Error Resume Next
   'MsgBox ActionName
'**** Release Report Test Set Count - Begin ***************
If ActionName = "RRTestSetCount" Then
 If User.IsInGroup("Custom Reports")Then
  Dim InputBox1
  Dim myFSO,WriteStuff,Stuff
  Set myFSO = CreateObject("Scripting.FileSystemObject")

  myFSO.DeleteFile "c:\ReleaseReport-TestSetCount.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestSetCount.txt", 8, True)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set Count"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you a count of Test Sets under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\TestSetCount Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
      InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestSetCount Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If

      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
              comCYCFS2.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & ""
              Set recsetCYCFS2 = comCYCFS2.Execute
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME") & ", " & recsetFSF1("CF_ITEM_NAME") &  ", " & recsetCYCFS2("CYCLECOUNT")
              WriteStuff.WriteLine(Stuff)
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " "
              Set recsetCYCFS1 = comCYCFS1.Execute
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CYCLECOUNT")
              WriteStuff.WriteLine(Stuff)

              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " "
              Set recsetCYCFS1 = comCYCFS1.Execute
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CYCLECOUNT")
              WriteStuff.WriteLine(Stuff)



              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & ""
            Set recsetCYCFS = comCYCFS.Execute
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CYCLECOUNT")
            WriteStuff.WriteLine(Stuff)
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " "
            Set recsetCYCFS = comCYCFS.Execute
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CYCLECOUNT")
            WriteStuff.WriteLine(Stuff)

           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " "
           Set recsetCYCF = comCYCF.Execute
           Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYCF("CYCLECOUNT")
           WriteStuff.WriteLine(Stuff)

          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select COUNT(*) as CYCLECOUNT from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " "
          Set recsetCYC = comCYC.Execute

          Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYC("CYCLECOUNT")
          WriteStuff.WriteLine(Stuff)

         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-TestSetCount.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestSetCount"
'**** Release Report Test Set Count - End ***************


'**** Release Report Test Set Details -Begin ***************
If ActionName = "RRTestSetDetails" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestSetDetails.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestSetDetails.txt", 8, True)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the details of the Test Sets under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\TestSet Details Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
      InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestSet Details Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If


      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
              comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME") & ", " & recsetFSF1("CF_ITEM_NAME") &  ", " & recsetCYCFS2("CY_CYCLE")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYCF("CY_CYCLE")
           WriteStuff.WriteLine(Stuff)
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYC("CY_CYCLE")
          WriteStuff.WriteLine(Stuff)
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestSetDetails.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestSetDetails"

'**** Release Report Test Set Details - End ***************


'**Release Report 3 ******
'**** Release Report Test Scenarios -Begin ***************
If ActionName = "RRTestSetScenarios" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestSetScenarios.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestSetScenarios.txt", 8, True)

  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set" & ", " & "ITG RequestId" & ", " & "Test Case" & ", " & "Test Instance" & ", " & "Execution Status"  & ", " & "Planned Tester" & ", " & "Planned Start Date" & ", " & "Planned Exec Date" & ", " & "Actual Exec Date" & ", " & "Actual Tester" & ", " & "Scripter" & ", " & "Business Requirements"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the details of the Test Scenarios for all the Test Sets under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\TestSet Scenarios Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
      InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestSet Scenarios Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If


      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
             comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist

              comTCTST1.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0

              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME") & "," & recsetFSF1("CF_ITEM_NAME") &  "," & recsetCYCFS2("CY_CYCLE") & "," & recsetTCTST1("TS_USER_02") & ","  & Replace(recsetTCTST1("TS_NAME"),",","") & ","  & recsetTCTST1("TC_TEST_INSTANCE")  & "," & recsetTCTST1("TC_STATUS") & "," & recsetTCTST1("TS_USER_11") & "," & recsetTCTST1("TS_USER_12") & "," & recsetTCTST1("TS_USER_13") & "," & recsetTCTST1("TC_EXEC_DATE") & "," & recsetTCTST1("TC_ACTUAL_TESTER") & "," & recsetTCTST1("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST1("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)

              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID,TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST2("TS_USER_02") & "," & Replace(recsetTCTST2("TS_NAME"),",","") & ","  & recsetTCTST2("TC_TEST_INSTANCE")  & "," & recsetTCTST2("TC_STATUS") & "," & recsetTCTST2("TS_USER_11") & "," & recsetTCTST2("TS_USER_12") & "," & recsetTCTST2("TS_USER_13") & "," & recsetTCTST2("TC_EXEC_DATE") & "," & recsetTCTST2("TC_ACTUAL_TESTER")  & "," & recsetTCTST2("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST2("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)
              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST3("TS_USER_02") & ","  & Replace(recsetTCTST3("TS_NAME"),",","") & ","  & recsetTCTST3("TC_TEST_INSTANCE")  & "," & recsetTCTST3("TC_STATUS") & "," & recsetTCTST3("TS_USER_11") & "," & recsetTCTST3("TS_USER_12") & "," & recsetTCTST3("TS_USER_13") & "," & recsetTCTST3("TC_EXEC_DATE") & "," & recsetTCTST3("TC_ACTUAL_TESTER") & "," & recsetTCTST3("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST3("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)
              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & ","  & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST4("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
            WriteStuff.WriteLine(Stuff)
            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST4("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
            WriteStuff.WriteLine(Stuff)
            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A" & "," & "N/A" & "," & recsetCYCF("CY_CYCLE") & "," & recsetTCTST5("TS_USER_02") & "," & Replace(recsetTCTST5("TS_NAME"),",","") & ","  & recsetTCTST5("TC_TEST_INSTANCE")  & "," & recsetTCTST5("TC_STATUS") & "," & recsetTCTST5("TS_USER_11") & "," & recsetTCTST5("TS_USER_12") & "," & recsetTCTST5("TS_USER_13") & "," & recsetTCTST5("TC_EXEC_DATE") & "," & recsetTCTST5("TC_ACTUAL_TESTER")& "," & recsetTCTST5("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST5("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
           WriteStuff.WriteLine(Stuff)
           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
          Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "," & recsetCYC("CY_CYCLE") & "," & recsetTCTST6("TS_USER_02") & "," & Replace(recsetTCTST6("TS_NAME"),",","") & ","  & recsetTCTST6("TC_TEST_INSTANCE")  & "," & recsetTCTST6("TC_STATUS") & "," & recsetTCTST6("TS_USER_11") & "," & recsetTCTST6("TS_USER_12") & "," & recsetTCTST6("TS_USER_13") & "," & recsetTCTST6("TC_EXEC_DATE") & "," & recsetTCTST6("TC_ACTUAL_TESTER") & "," & recsetTCTST6("TS_USER_14") & "," & Replace(Replace(Replace(recsetTCTST6("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
          WriteStuff.WriteLine(Stuff)
          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestSetScenarios.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestSetDetails"

'**** Release Report Test Scenarios - End ***************





'**Release Report 4 ******
'**** Release Report Test Set Runs -Begin ***************
If ActionName = "RRTestScenariosCount" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestScenariosCount.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestScenariosCount.txt", 8, True)
   Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set" & ", " & "Status-Passed" & ", "  & "Status-Failed" & ", " & "Status-NA" & ", " & "Status-NoRun" & ", " & "Status-NotCompleted"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("Enter the Release Number. This report will give you a count of Test Scenarios by TestSet and Status","Release\TestScenariosCount Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command

   'Run table command sets
   Set comRun1 = tdc.command
   Set comRun2 = tdc.command
   Set comRun3 = tdc.command
   Set comRun4 = tdc.command

   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then

     InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestScenarios Count Report")
     If len(InputBox2) <= 0 Then
     comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
     Else
     comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
     End If

     Set recsetM = comM.Execute
     WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
              comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist
              STPass = 0
              STFail = 0
              STNA = 0
              STNC = 0
              STNR = 0
              comTCTST1.CommandText = "SELECT TC_CYCLE_ID, TC_STATUS  FROM TESTCYCL  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & ""
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0

              Select Case recsetTCTST1("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1

              End Select
              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME") & ", " & recsetFSF1("CF_ITEM_NAME") &  ", " & recsetCYCFS2("CY_CYCLE") & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              STPass = 0
              STFail = 0
              STNA = 0
              STNC = 0
              STNR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID, TC_STATUS FROM TESTCYCL  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
              Select Case recsetTCTST2("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select
              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
               Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE")  & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              STPass = 0
              STFail = 0
              STNA = 0
              STNC = 0
              STNR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_STATUS  FROM TESTCYCL  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              Select Case recsetTCTST3("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select

              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE")  & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
              WriteStuff.WriteLine(Stuff)
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
              STPass = 0
              STFail = 0
              STNA = 0
              STNC = 0
              STNR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID,TC_STATUS  FROM TESTCYCL  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & ""
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
             Select Case recsetTCTST4("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select

            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE")  & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
            WriteStuff.WriteLine(Stuff)
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
             STPass = 0
              STFail = 0
              STNA = 0
              STNC = 0
              STNR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID,TC_STATUS FROM TESTCYCL WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            Select Case recsetTCTST4("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select

            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE") & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
            WriteStuff.WriteLine(Stuff)
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           STPass = 0
           STFail = 0
           STNA = 0
           STNC = 0
           STNR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID,TC_STATUS FROM TESTCYCL  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
           Select Case recsetTCTST5("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select

           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYCF("CY_CYCLE") & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
           WriteStuff.WriteLine(Stuff)
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          STPass = 0
           STFail = 0
           STNA = 0
           STNC = 0
           STNR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID,TC_STATUS  FROM TESTCYCL WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
           Select Case recsetTCTST6("TC_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select

          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYC("CY_CYCLE") & ", " & STPass & ", " & STFail & ", " & STNA & ", " & STNR & ", " & STNC
          WriteStuff.WriteLine(Stuff)
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestScenariosCount.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestScenariosCount"
'**** Release Report TestScenariosCount - End ***************

'**Release Report 5 ******
'**** Release Report Test Run Count -Begin ***************
If ActionName = "RRTestRunsCount" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestRunsCount.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestRunsCount.txt", 8, True)

  Stuff = "Level 1" & "," & "Level 2" & "," & "Level 3"  & "," & "Level 4" & "," & "Level 5" & "," & "Level 6"  & "," & "Level 7"  & "," & "Test Set" & "," & "ITG RequestId" & "," & "Test Case" & "," & "Test Instance" & "," & "Execution Status"  & ","  & "Actual Exec Date" & "," & "Actual Tester" & "," & "Scripter" & "," & "Runs-Passed" & "," & "Runs-Failed" & ","  & "Runs-NA" & "," & "Runs-NC" & "," & "Runs-NR"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the count of Test Runs for all the Test Scenarios under each Test Set for a Release by Release Number & HHSC Sub folders. Enter the Release Number","Release\TestSet Run Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command

   'Run table command sets
   Set comRun1 = tdc.command
   Set comRun2 = tdc.command
   Set comRun3 = tdc.command
   Set comRun4 = tdc.command
   Set comRun5 = tdc.command
   Set comRun6 = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
     InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestSet Run Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If

      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
             comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist

              comTCTST1.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0
              STPass = 0
              STFail = 0
              STNR = 0
              STNA = 0
              STNC = 0

              comRun1.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST1("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST1("TC_TESTCYCL_ID") & " "
              Set recsetRun1 = comRun1.Execute
              If recsetRun1.EOR = 0 Then
              WHILE recsetRun1.EOR = 0
               Select Case recsetRun1("RN_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select
              recsetRun1.Next
              WEND
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME") & "," & recsetFSF1("CF_ITEM_NAME") &  "," & recsetCYCFS2("CY_CYCLE") & "," & recsetTCTST1("TS_USER_02") & "," & Replace(recsetTCTST1("TS_NAME"),",","") & ","  & recsetTCTST1("TC_TEST_INSTANCE")  & "," & recsetTCTST1("TC_STATUS") &  "," & recsetTCTST1("TC_EXEC_DATE") & "," & recsetTCTST1("TC_ACTUAL_TESTER") & "," & recsetTCTST1("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
              WriteStuff.WriteLine(Stuff)
              End If

              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID,TS_USER_02,TS_USER_14, TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
              STPass = 0
              STFail = 0
              STNR = 0
              STNA = 0
              STNC = 0

              comRun2.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST2("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST2("TC_TESTCYCL_ID") & " "
              Set recsetRun2 = comRun2.Execute
              If recsetRun2.EOR = 0 Then
              WHILE recsetRun2.EOR = 0
               Select Case recsetRun2("RN_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select
              recsetRun2.Next
              WEND

             ' Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE") & ", " & recsetTCTST2("TS_NAME") & ", "  & recsetTCTST2("TC_TEST_INSTANCE")  & ", " & recsetTCTST2("TC_STATUS") & ", " & recsetTCTST2("TC_TESTER_NAME") & ", " & recsetTCTST2("TC_USER_03") & "," & recsetTCTST2("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST2("TC_EXEC_DATE") & ", " & recsetTCTST2("TC_ACTUAL_TESTER")
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST2("TS_USER_02") & "," & Replace(recsetTCTST2("TS_NAME"),",","") & ","  & recsetTCTST2("TC_TEST_INSTANCE")  & "," & recsetTCTST2("TC_STATUS") &  "," & recsetTCTST2("TC_EXEC_DATE") & "," & recsetTCTST2("TC_ACTUAL_TESTER") & "," & recsetTCTST2("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
              WriteStuff.WriteLine(Stuff)
              End If
              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_USER_14,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              STPass = 0
              STFail = 0
              STNR = 0
              STNA = 0
              STNC = 0

              comRun3.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST3("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST3("TC_TESTCYCL_ID") & " "
              Set recsetRun3 = comRun3.Execute
              If recsetRun3.EOR = 0 Then
              WHILE recsetRun3.EOR = 0
               Select Case recsetRun3("RN_STATUS")
              Case "Passed"
               STPass = STPass + 1
              Case "Failed"
               STFail = STFail + 1
              Case "No Run"
               STNR = STNR + 1
              Case "N/A"
               STNA = STNA + 1
              Case "Not Completed"
               STNC = STNC + 1
              End Select
              recsetRun3.Next
              WEND
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST3("TS_USER_02") & ","  & Replace(recsetTCTST3("TS_NAME"),",","") & ","  & recsetTCTST3("TC_TEST_INSTANCE")  & "," & recsetTCTST3("TC_STATUS") &  "," & recsetTCTST3("TC_EXEC_DATE") & "," & recsetTCTST3("TC_ACTUAL_TESTER") & "," & recsetTCTST3("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
              WriteStuff.WriteLine(Stuff)
              End If
              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_USER_14,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
             STPass = 0
             STFail = 0
             STNR = 0
             STNA = 0
             STNC = 0

             comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "
             Set recsetRun4 = comRun4.Execute
             If recsetRun4.EOR = 0 Then
             WHILE recsetRun4.EOR = 0
             Select Case recsetRun4("RN_STATUS")
             Case "Passed"
               STPass = STPass + 1
             Case "Failed"
               STFail = STFail + 1
             Case "No Run"
               STNR = STNR + 1
             Case "N/A"
               STNA = STNA + 1
             Case "Not Completed"
               STNC = STNC + 1
             End Select
              recsetRun4.Next
             WEND
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") &  "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
            WriteStuff.WriteLine(Stuff)
            End If
            Set recsetRun4 = NOTHING
            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_USER_14,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
             STPass = 0
             STFail = 0
             STNR = 0
             STNA = 0
             STNC = 0

             comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "
             Set recsetRun4 = comRun4.Execute
             If recsetRun4.EOR = 0 Then
             WHILE recsetRun4.EOR = 0
             Select Case recsetRun4("RN_STATUS")
             Case "Passed"
               STPass = STPass + 1
             Case "Failed"
               STFail = STFail + 1
             Case "No Run"
               STNR = STNR + 1
             Case "N/A"
               STNA = STNA + 1
             Case "Not Completed"
               STNC = STNC + 1
             End Select
              recsetRun4.Next
             WEND
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") &  "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
            WriteStuff.WriteLine(Stuff)
            End If
            Set recsetRun4 = Nothing
            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0

            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_USER_14,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
            STPass = 0
             STFail = 0
             STNR = 0
             STNA = 0
             STNC = 0

             comRun5.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST5("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST5("TC_TESTCYCL_ID") & " "
             Set recsetRun5 = comRun5.Execute
             If recsetRun5.EOR = 0 Then
             WHILE recsetRun5.EOR = 0
             Select Case recsetRun5("RN_STATUS")
             Case "Passed"
               STPass = STPass + 1
             Case "Failed"
               STFail = STFail + 1
             Case "No Run"
               STNR = STNR + 1
             Case "N/A"
               STNA = STNA + 1
             Case "Not Completed"
               STNC = STNC + 1
             End Select
              recsetRun5.Next
             WEND
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A" & "," & "N/A" & "," & recsetCYCF("CY_CYCLE") & "," & recsetTCTST5("TS_USER_02") & "," & Replace(recsetTCTST5("TS_NAME"),",","") & ","  & recsetTCTST5("TC_TEST_INSTANCE")  & "," & recsetTCTST5("TC_STATUS") &  "," & recsetTCTST5("TC_EXEC_DATE") & "," & recsetTCTST5("TC_ACTUAL_TESTER") & "," & recsetTCTST5("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
           WriteStuff.WriteLine(Stuff)
           End If
           Set recsetRun5 = NOTHING
           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_USER_14,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
           STPass = 0
           STFail = 0
           STNR = 0
           STNA = 0
           STNC = 0

           comRun6.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST6("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST6("TC_TESTCYCL_ID") & " "
             Set recsetRun6 = comRun6.Execute
             If recsetRun6.EOR = 0 Then
             WHILE recsetRun6.EOR = 0
             Select Case recsetRun6("RN_STATUS")
             Case "Passed"
               STPass = STPass + 1
             Case "Failed"
               STFail = STFail + 1
             Case "No Run"
               STNR = STNR + 1
             Case "N/A"
               STNA = STNA + 1
             Case "Not Completed"
               STNC = STNC + 1
             End Select
              recsetRun6.Next
             WEND
          Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "," & recsetCYC("CY_CYCLE") & "," & recsetTCTST6("TS_USER_02") & "," & Replace(recsetTCTST6("TS_NAME"),",","") & ","  & recsetTCTST6("TC_TEST_INSTANCE")  & "," & recsetTCTST6("TC_STATUS") &  "," & recsetTCTST6("TC_EXEC_DATE") & "," & recsetTCTST6("TC_ACTUAL_TESTER") & "," & recsetTCTST6("TS_USER_14") & "," & STPass & "," & STFail & "," & STNA & "," & STNC & "," & STNR
          WriteStuff.WriteLine(Stuff)
          End If
          Set recsetRun6 = NOTHING
          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        'WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestRunsCount.txt.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
   Set recsetRun1  = nothing
   Set recsetRun2  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RRTestRunsCount"

'**** Release Report Test Runs Count - End ***************


'**** Release Report TestRunDetails - Begin ***************
If ActionName = "RRTestRunDetails" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestRunDetails.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestRunDetails.txt", 8, True)

  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set" & ", " & "ITG RequestId" & ", " & "Test Case" & ", " & "Test Instance" & ", " & "Execution Status"  & ", " & "Planned Tester" & ", " & "Planned Start Date" & ", " & "Planned Exec Date" & ", " & "Actual Exec Date" & ", " & "Actual Tester" & ", " & "Scripter"  & ", " & "Run Name" & ", " & "Run Tester" & ", " & "Run Status"  & ", " & "Run Execution Date"  & ", " & "Client IDs" & ", " & "EDG IDs"   & ", " & "CaseIDs"   & ", " & "Business Requirement"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the details of the Test Runs for all the Test Sets under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\TestSet Scenarios Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command
   Set comRun1  = tdc.command
   Set comRun2  = tdc.command
   Set comRun3  = tdc.command
    Set comRun4  = tdc.command
   Set comRun5  = tdc.command
   Set comRun6  = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
     InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\TestSet Scenarios Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If


      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME in ('UAT','SIT') and CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
             comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist

              comTCTST1.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0
               'Check to see the runs

              comRun1.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST1("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST1("TC_TESTCYCL_ID") & " "
              Set recsetRun1 = comRun1.Execute
              WHILE recsetRun1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME") & "," & recsetFSF1("CF_ITEM_NAME") &  "," & recsetCYCFS2("CY_CYCLE") & "," & recsetTCTST1("TS_USER_02") & ","  & Replace(recsetTCTST1("TS_NAME"),",","") & ","  & recsetTCTST1("TC_TEST_INSTANCE")  & "," & recsetTCTST1("TC_STATUS") & "," & recsetTCTST1("TS_USER_11") & "," & recsetTCTST1("TS_USER_12") & "," & recsetTCTST1("TS_USER_13") & "," & recsetTCTST1("TC_EXEC_DATE") & "," & recsetTCTST1("TC_ACTUAL_TESTER") &  "," & recsetTCTST1("TS_USER_14") & "," & recsetRun1("RN_RUN_NAME") & "," &  recsetRun1("RN_TESTER_NAME") & "," &  recsetRun1("RN_STATUS") & ","  &  recsetRun1("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun1("RN_USER_07"),",","/") & ","  &  Replace(recsetRun1("RN_USER_06"),",","/") & ","  &  Replace(recsetRun1("RN_USER_05"),",","/") & ","  & Replace(Replace(Replace(recsetTCTST1("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)
              recsetRun1.Next
              WEND

               '***Runs End****


              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
               'Check to see the runs

              comRun2.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST2("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST2("TC_TESTCYCL_ID") & " "
              Set recsetRun2 = comRun2.Execute
              WHILE recsetRun2.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST2("TS_USER_02") & ","  & Replace(recsetTCTST2("TS_NAME"),",","") & ","  & recsetTCTST2("TC_TEST_INSTANCE")  & "," & recsetTCTST2("TC_STATUS") & "," & recsetTCTST2("TS_USER_11") & "," & recsetTCTST2("TS_USER_12") & "," & recsetTCTST2("TS_USER_13") & "," & recsetTCTST2("TC_EXEC_DATE") & "," & recsetTCTST2("TC_ACTUAL_TESTER") & "," & recsetTCTST2("TS_USER_14") & "," & recsetRun2("RN_RUN_NAME") & "," &  recsetRun2("RN_TESTER_NAME") & "," &  recsetRun2("RN_STATUS") & ","  &  recsetRun2("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun2("RN_USER_07"),",","/") & ","  &  Replace(recsetRun2("RN_USER_06"),",","/") & ","  &  Replace(recsetRun2("RN_USER_05"),",","/") & ","  & Replace(Replace(Replace(recsetTCTST2("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)
              recsetRun2.Next
              WEND

               '***Runs End****

              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              comRun3.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST3("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST3("TC_TESTCYCL_ID") & " "

              Set recsetRun3 = comRun3.Execute
              WHILE recsetRun3.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST3("TS_USER_02") & "," & Replace(recsetTCTST3("TS_NAME"),",","") & ","  & recsetTCTST3("TC_TEST_INSTANCE")  & "," & recsetTCTST3("TC_STATUS") & "," & recsetTCTST3("TS_USER_11") & "," & recsetTCTST3("TS_USER_12") & "," & recsetTCTST3("TS_USER_13") & "," & recsetTCTST3("TC_EXEC_DATE") & "," & recsetTCTST3("TC_ACTUAL_TESTER") & "," & recsetTCTST3("TS_USER_14") & "," & recsetRun3("RN_RUN_NAME") & "," &  recsetRun3("RN_TESTER_NAME") & "," &  recsetRun3("RN_STATUS") & ","  &  recsetRun3("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun3("RN_USER_07"),",","/") & ","  &  Replace(recsetRun3("RN_USER_06"),",","/") & ","  &  Replace(recsetRun3("RN_USER_05"),",","/") & "," & Replace(Replace(Replace(recsetTCTST3("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
              WriteStuff.WriteLine(Stuff)
              recsetRun3.Next
              WEND

              '***Run End
              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TESTCYCL_ID,TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "

            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & recsetRun4("RN_RUN_NAME") & "," &  recsetRun4("RN_TESTER_NAME") & "," &  recsetRun4("RN_STATUS") & ","  &  recsetRun4("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun4("RN_USER_07"),",","/") & ","  &  Replace(recsetRun4("RN_USER_06"),",","/") & ","  &  Replace(recsetRun4("RN_USER_05"),",","/") & "," & Replace(Replace(Replace(recsetTCTST4("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
            WriteStuff.WriteLine(Stuff)
            recsetRun4.Next
            WEND
            Set recsetRun4 = NOTHING
              '***Run End

            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "
            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0

            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & recsetRun4("RN_RUN_NAME") & "," &  recsetRun4("RN_TESTER_NAME") & "," &  recsetRun4("RN_STATUS") & ","  &  recsetRun4("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun4("RN_USER_07"),",","/") & ","  &  Replace(recsetRun4("RN_USER_06"),",","/") & ","  &  Replace(recsetRun4("RN_USER_05"),",","/") & "," & Replace(Replace(Replace(recsetTCTST4("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
            WriteStuff.WriteLine(Stuff)
            recsetRun4.Next
            WEND


            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
           comRun5.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST5("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST5("TC_TESTCYCL_ID") & " "
           Set recsetRun5 = comRun5.Execute
           WHILE recsetRun5.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A" & "," & "N/A" & "," & recsetCYCF("CY_CYCLE") & "," & recsetTCTST5("TS_USER_02") & ","  & Replace(recsetTCTST5("TS_NAME"),",","") & ","  & recsetTCTST5("TC_TEST_INSTANCE")  & "," & recsetTCTST5("TC_STATUS") & "," & recsetTCTST5("TS_USER_11") & "," & recsetTCTST5("TS_USER_12") & "," & recsetTCTST5("TS_USER_13") & "," & recsetTCTST5("TC_EXEC_DATE") & "," & recsetTCTST5("TC_ACTUAL_TESTER") & "," & recsetTCTST5("TS_USER_14") & "," & recsetRun5("RN_RUN_NAME") & "," &  recsetRun5("RN_TESTER_NAME") & "," &  recsetRun5("RN_STATUS") & ","  &  recsetRun5("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun5("RN_USER_07"),",","/") & ","  &  Replace(recsetRun5("RN_USER_06"),",","/") & ","  &  Replace(recsetRun5("RN_USER_05"),",","/") & ","  & Replace(Replace(Replace(recsetTCTST5("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
           WriteStuff.WriteLine(Stuff)
           recsetRun5.Next
           WEND

           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TS_USER_15,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
          comRun6.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST6("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST6("TC_TESTCYCL_ID") & " "
           Set recsetRun6 = comRun6.Execute
           WHILE recsetRun6.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "," & recsetCYC("CY_CYCLE") & "," & recsetTCTST6("TS_USER_02") & "," & Replace(recsetTCTST6("TS_NAME"),",","") & ","  & recsetTCTST6("TC_TEST_INSTANCE")  & "," & recsetTCTST6("TC_STATUS") & "," & recsetTCTST6("TS_USER_11") & "," & recsetTCTST6("TS_USER_12") & "," & recsetTCTST6("TS_USER_13") & "," & recsetTCTST6("TC_EXEC_DATE") & "," & recsetTCTST6("TC_ACTUAL_TESTER") & "," & recsetTCTST6("TS_USER_14") & "," & recsetRun6("RN_RUN_NAME") & "," &  recsetRun6("RN_TESTER_NAME") & "," &  recsetRun6("RN_STATUS") & ","  &  recsetRun6("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun6("RN_USER_07"),",","/") & ","  &  Replace(recsetRun6("RN_USER_06"),",","/") & ","  &  Replace(recsetRun6("RN_USER_05"),",","/") & "," & Replace(Replace(Replace(recsetTCTST6("TS_USER_15"),",",""),Chr(13)," "),Chr(10)," ")
           WriteStuff.WriteLine(Stuff)
           recsetRun6.Next
           WEND


          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "N/A"  & "," & "N/A" & "," & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestRunDetails.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
     Set recsetRun1  = nothing
    Set recsetRun2  = nothing
     Set recsetRun3  = nothing
     Set recsetRun4  = nothing
    Set recsetRun5  = nothing
     Set recsetRun6  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestRunDetails"

'**** Release Report TestRunDetails - End *****************




'****SIT Test Runs Details ************************
'***********************************************

'**** Release Report TestRunDetails - Begin ***************
If ActionName = "SITTestRunDetails" Then
 If User.IsInGroup("DLCustom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\SIT_TestRunDetails.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\SIT_TestRunDetails.txt", 8, True)

  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set" & ", " & "Test Case" & ", " & "Test Instance" & ", " & "Execution Status"  & ", " & "Planned Tester" & ", " & "Planned Start Date" & ", " & "Planned Exec Date" & ", " & "Actual Exec Date" & ", " & "Actual Tester"  & ", " & "Run Name" & ", " & "Run Tester" & ", " & "Run Status"  & ", " & "Run Execution Date"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the details of the Test Runs for all the Test Sets under each Release Folder by Release Number & Deloitte Sub folders. Enter the Release Number","Release\TestSet Scenarios Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command
   Set comRun1  = tdc.command
   Set comRun2  = tdc.command
   Set comRun3  = tdc.command
    Set comRun4  = tdc.command
   Set comRun5  = tdc.command
   Set comRun6  = tdc.command
   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'Deloitte' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME in ('Maint SRs','Mod SRs','Regression','Smoke Tests') and CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
             comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist

              comTCTST1.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0
               'Check to see the runs

              comRun1.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST1("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST1("TC_TESTCYCL_ID") & " "
              Set recsetRun1 = comRun1.Execute
              WHILE recsetRun1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME") & ", " & recsetFSF1("CF_ITEM_NAME") &  ", " & recsetCYCFS2("CY_CYCLE") & ", " & Replace(recsetTCTST1("TS_NAME"),",","") & ", "  & recsetTCTST1("TC_TEST_INSTANCE")  & ", " & recsetTCTST1("TC_STATUS") & ", " & recsetTCTST1("TC_TESTER_NAME") & ", " & recsetTCTST1("TC_USER_03") & "," & recsetTCTST1("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST1("TC_EXEC_DATE") & ", " & recsetTCTST1("TC_ACTUAL_TESTER") & ", " & recsetRun1("RN_RUN_NAME") & ", " &  recsetRun1("RN_TESTER_NAME") & ", " &  recsetRun1("RN_STATUS") & ", "  &  recsetRun1("RN_EXECUTION_DATE")
              WriteStuff.WriteLine(Stuff)
              recsetRun1.Next
              WEND



               '***Runs End****



              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
               'Check to see the runs

              comRun2.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST2("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST2("TC_TESTCYCL_ID") & " "
              Set recsetRun2 = comRun2.Execute
              WHILE recsetRun2.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE") & ", " & Replace(recsetTCTST2("TS_NAME"),",","") & ", "  & recsetTCTST2("TC_TEST_INSTANCE")  & ", " & recsetTCTST2("TC_STATUS") & ", " & recsetTCTST2("TC_TESTER_NAME") & ", " & recsetTCTST2("TC_USER_03") & "," & recsetTCTST2("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST2("TC_EXEC_DATE") & ", " & recsetTCTST2("TC_ACTUAL_TESTER")& ", " & recsetRun2("RN_RUN_NAME") & ", " &  recsetRun2("RN_TESTER_NAME") & ", " &  recsetRun2("RN_STATUS") & ", "  &  recsetRun2("RN_EXECUTION_DATE")
              WriteStuff.WriteLine(Stuff)
              recsetRun2.Next
              WEND



               '***Runs End****

              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              comRun3.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST3("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST3("TC_TESTCYCL_ID") & " "

              Set recsetRun3 = comRun3.Execute
              WHILE recsetRun3.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", " & recsetFSF("CF_ITEM_NAME")& ", "  & "N/A"  &  ", " & recsetCYCFS1("CY_CYCLE") & ", " & Replace(recsetTCTST3("TS_NAME"),",","") & ", "  & recsetTCTST3("TC_TEST_INSTANCE")  & ", " & recsetTCTST3("TC_STATUS") & ", " & recsetTCTST3("TC_TESTER_NAME") & ", " & recsetTCTST3("TC_USER_03") & "," & recsetTCTST3("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST3("TC_EXEC_DATE") & ", " & recsetTCTST3("TC_ACTUAL_TESTER")& ", " & recsetRun3("RN_RUN_NAME") & ", " &  recsetRun3("RN_TESTER_NAME") & ", " &  recsetRun3("RN_STATUS") & ", "  &  recsetRun3("RN_EXECUTION_DATE")
              WriteStuff.WriteLine(Stuff)
              recsetRun3.Next
              WEND

              '***Run End
              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TESTCYCL_ID,TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "

            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE") & ", " & Replace(recsetTCTST4("TS_NAME"),",","") & ", "  & recsetTCTST4("TC_TEST_INSTANCE")  & ", " & recsetTCTST4("TC_STATUS") & ", " & recsetTCTST4("TC_TESTER_NAME") & ", " & recsetTCTST4("TC_USER_03") & "," & recsetTCTST4("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST4("TC_EXEC_DATE") & ", " & recsetTCTST4("TC_ACTUAL_TESTER")& ", " & recsetRun4("RN_RUN_NAME") & ", " &  recsetRun4("RN_TESTER_NAME") & ", " &  recsetRun4("RN_STATUS") & ", "  &  recsetRun4("RN_EXECUTION_DATE")
            WriteStuff.WriteLine(Stuff)
            recsetRun4.Next
            WEND
            Set recsetRun4 = NOTHING
              '***Run End

            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "
            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0

            Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & recsetFS("CF_ITEM_NAME") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("CY_CYCLE") & ", " & Replace(recsetTCTST4("TS_NAME"),",","") & ", "  & recsetTCTST4("TC_TEST_INSTANCE")  & ", " & recsetTCTST4("TC_STATUS") & ", " & recsetTCTST4("TC_TESTER_NAME") & ", " & recsetTCTST4("TC_USER_03") & "," & recsetTCTST4("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST4("TC_EXEC_DATE") & ", " & recsetTCTST4("TC_ACTUAL_TESTER") & ", " & recsetRun4("RN_RUN_NAME") & ", " &  recsetRun4("RN_TESTER_NAME") & ", " &  recsetRun4("RN_STATUS") & ", "  &  recsetRun4("RN_EXECUTION_DATE")
            WriteStuff.WriteLine(Stuff)
            recsetRun4.Next
            WEND


            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
           comRun5.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST5("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST5("TC_TESTCYCL_ID") & " "
           Set recsetRun5 = comRun5.Execute
           WHILE recsetRun5.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME")  & ", " & recsetF("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYCF("CY_CYCLE") & ", " & Replace(recsetTCTST5("TS_NAME"),",","") & ", "  & recsetTCTST5("TC_TEST_INSTANCE")  & ", " & recsetTCTST5("TC_STATUS") & ", " & recsetTCTST5("TC_TESTER_NAME") & ", " & recsetTCTST5("TC_USER_03") & "," & recsetTCTST5("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST5("TC_EXEC_DATE") & ", " & recsetTCTST5("TC_ACTUAL_TESTER") & ", " & recsetRun5("RN_RUN_NAME") & ", " &  recsetRun5("RN_TESTER_NAME") & ", " &  recsetRun5("RN_STATUS") & ", "  &  recsetRun5("RN_EXECUTION_DATE")
           WriteStuff.WriteLine(Stuff)
           recsetRun5.Next
           WEND

           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_NAME,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
          comRun6.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST6("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST6("TC_TESTCYCL_ID") & " "
           Set recsetRun6 = comRun6.Execute
           WHILE recsetRun6.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & recsetS("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A" & ", " & recsetCYC("CY_CYCLE") & ", " & Replace(recsetTCTST6("TS_NAME"),",","") & ", "  & recsetTCTST6("TC_TEST_INSTANCE")  & ", " & recsetTCTST6("TC_STATUS") & ", " & recsetTCTST6("TC_TESTER_NAME") & ", " & recsetTCTST6("TC_USER_03") & "," & recsetTCTST6("TC_PLAN_SCHEDULING_DATE") & ", " & recsetTCTST6("TC_EXEC_DATE") & ", " & recsetTCTST6("TC_ACTUAL_TESTER") & ", " & recsetRun6("RN_RUN_NAME") & ", " &  recsetRun6("RN_TESTER_NAME") & ", " &  recsetRun6("RN_STATUS") & ", "  &  recsetRun6("RN_EXECUTION_DATE")
           WriteStuff.WriteLine(Stuff)
           recsetRun6.Next
           WEND


          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME") & ", " & "N/A" & ", " & "N/A"  & ", " & "N/A" & ", " & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\SIT_TestRunDetails.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
     Set recsetRun1  = nothing
    Set recsetRun2  = nothing
     Set recsetRun3  = nothing
     Set recsetRun4  = nothing
    Set recsetRun5  = nothing
     Set recsetRun6  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "SIT_TestRunDetails"

'**** Release Report SIT_TestRunDetails - End *****************
'***********************************************

'STEP DETAILS
'**** Release Report TestRunStepDetails - Begin ***************
If ActionName = "RRTestRunStepDetails" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\ReleaseReport-TestRunStepDetails.txt"

  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestRunStepDetails.txt", 8, True)

  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Set" & ", " & "ITG RequestId" & ", " & "Test Case" & ", " & "Test Instance" & ", " & "Execution Status"  & ", " & "Planned Tester" & ", " & "Planned Start Date" & ", " & "Planned Exec Date" & ", " & "Actual Exec Date" & ", " & "Actual Tester" & ", " & "Scripter"  & ", " & "Run Name" & ", " & "Run Tester" & ", " & "Run Status"  & ", " & "Run Execution Date"  & ", " & "Client IDs" & ", " & "EDG IDs"   & ", " & "CaseIDs"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the count of Steps of the Test Runs for all the Test Sets under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\Run Step Count Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command

   'TestCycle,Test table command sets
   Set comTCTST1 = tdc.command
   Set comTCTST2 = tdc.command
   Set comTCTST3 = tdc.command
   Set comTCTST4 = tdc.command
   Set comTCTST5 = tdc.command
   Set comTCTST6 = tdc.command
   Set comRun1  = tdc.command
   Set comRun2  = tdc.command
   Set comRun3  = tdc.command
   Set comRun4  = tdc.command
   Set comRun5  = tdc.command
   Set comRun6  = tdc.command
   Set comStep1  = tdc.command
   Set comStep2  = tdc.command
   Set comStep3  = tdc.command
   Set comStep4  = tdc.command
   Set comStep5  = tdc.command
   Set comStep6  = tdc.command

   'Main Process Logic
   comR.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     If recsetR("CF_NO_OF_SONS") > 0 Then
      InputBox2 = InputBox ("Enter the name of the Sub Folder to report on.Leave the field blank if you want to report on all the sub folders","Release\Run Step Count Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      Else
      comM.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = '" & InputBox2 & "' and CF_FATHER_ID = '" & recsetR("CF_ITEM_ID") & "' Order by CF_ITEM_NAME"
      End If

      Set recsetM = comM.Execute
      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       If recsetM("CF_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME = 'HHSC' and CF_FATHER_ID = " & recsetM("CF_ITEM_ID")
        Set recsetS = comS.Execute
        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         If recsetS("CF_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_ITEM_NAME in ('UAT','SIT') and CF_FATHER_ID = " & recsetS("CF_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          If recsetF("CF_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetF("CF_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           If recsetFS("CF_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFS("CF_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              If recsetFSF("CF_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select CF_ITEM_ID,CF_FATHER_ID,CF_ITEM_NAME,CF_NO_OF_SONS from CYCL_FOLD where CF_FATHER_ID = " & recsetFSF("CF_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the CYCLE table if any
             comCYCFS2.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF1("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
               'Check to see if test scenarios exist

              comTCTST1.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS2("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST1 = comTCTST1.Execute
              WHILE recsetTCTST1.EOR = 0
               'Check to see the runs

              comRun1.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST1("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST1("TC_TESTCYCL_ID") & " "
              Set recsetRun1 = comRun1.Execute
              WHILE recsetRun1.EOR = 0
              comStep1.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun1("RN_RUN_ID")
              Set recsetStep1 = comStep1.Execute
              WHILE recsetStep1.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME") & "," & recsetFSF1("CF_ITEM_NAME") &  "," & recsetCYCFS2("CY_CYCLE") & "," & recsetTCTST1("TS_USER_02") & ","  & Replace(recsetTCTST1("TS_NAME"),",","") & ","  & recsetTCTST1("TC_TEST_INSTANCE")  & "," & recsetTCTST1("TC_STATUS") & "," & recsetTCTST1("TS_USER_11") & "," & recsetTCTST1("TS_USER_12") & "," & recsetTCTST1("TS_USER_13") & "," & recsetTCTST1("TC_EXEC_DATE") & "," & recsetTCTST1("TC_ACTUAL_TESTER") &  "," & recsetTCTST1("TS_USER_14") & "," & recsetRun1("RN_RUN_NAME") & "," &  recsetRun1("RN_TESTER_NAME") & "," &  recsetRun1("RN_STATUS") & ","  &  recsetRun1("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun1("RN_USER_07"),",","/") & ","  &  Replace(recsetRun1("RN_USER_06"),",","/") & ","  &  Replace(recsetRun1("RN_USER_05"),",","/") & "," & recsetStep1("ST_STEP_NAME") & "," & recsetStep1("ST_STATUS")
              WriteStuff.WriteLine(Stuff)
              recsetStep1.Next
              WEND
              recsetRun1.Next
              WEND

               '***Runs End****


              recsetTCTST1.Next
              WEND 'recsetTCTST1.EOR = 0
              recsetCYCFS2.Next
              WEND 'recsetFSF2.EOR = 0
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST2.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST2 = comTCTST2.Execute
              WHILE recsetTCTST2.EOR = 0
               'Check to see the runs

              comRun2.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST2("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST2("TC_TESTCYCL_ID") & " "
              Set recsetRun2 = comRun2.Execute
              WHILE recsetRun2.EOR = 0
              comStep2.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun2("RN_RUN_ID")
              Set recsetStep2 = comStep2.Execute
              WHILE recsetStep2.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST2("TS_USER_02") & ","  & Replace(recsetTCTST2("TS_NAME"),",","") & ","  & recsetTCTST2("TC_TEST_INSTANCE")  & "," & recsetTCTST2("TC_STATUS") & "," & recsetTCTST2("TS_USER_11") & "," & recsetTCTST2("TS_USER_12") & "," & recsetTCTST2("TS_USER_13") & "," & recsetTCTST2("TC_EXEC_DATE") & "," & recsetTCTST2("TC_ACTUAL_TESTER") & "," & recsetTCTST2("TS_USER_14") & "," & recsetRun2("RN_RUN_NAME") & "," &  recsetRun2("RN_TESTER_NAME") & "," &  recsetRun2("RN_STATUS") & ","  &  recsetRun2("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun2("RN_USER_07"),",","/") & ","  &  Replace(recsetRun2("RN_USER_06"),",","/") & ","  &  Replace(recsetRun2("RN_USER_05"),",","/") & "," & recsetStep2("ST_STEP_NAME") & "," & recsetStep2("ST_STATUS")
              WriteStuff.WriteLine(Stuff)
              recsetStep2.Next
              Wend
              recsetRun2.Next
              WEND

               '***Runs End****

              recsetTCTST2.Next
              WEND 'recsetTCTST2.EOR = 0
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0
              Set recsetCYCFS1 = Nothing

              '*****

              Else 'recsetFSF("CF_NO_OF_SONS") > 0
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFSF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comTCTST3.CommandText = "SELECT TC_CYCLE_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER,TC_TESTCYCL_ID  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS1("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
              Set recsetTCTST3 = comTCTST3.Execute
              WHILE recsetTCTST3.EOR = 0
              comRun3.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST3("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST3("TC_TESTCYCL_ID") & " "

              Set recsetRun3 = comRun3.Execute
              WHILE recsetRun3.EOR = 0
              comStep3.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun3("RN_RUN_ID")
              Set recsetStep3 = comStep3.Execute
              WHILE recsetStep3.EOR = 0
              Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & "," & recsetFSF("CF_ITEM_NAME")& ","  & "N/A"  &  "," & recsetCYCFS1("CY_CYCLE") & "," & recsetTCTST3("TS_USER_02") & "," & Replace(recsetTCTST3("TS_NAME"),",","") & ","  & recsetTCTST3("TC_TEST_INSTANCE")  & "," & recsetTCTST3("TC_STATUS") & "," & recsetTCTST3("TS_USER_11") & "," & recsetTCTST3("TS_USER_12") & "," & recsetTCTST3("TS_USER_13") & "," & recsetTCTST3("TC_EXEC_DATE") & "," & recsetTCTST3("TC_ACTUAL_TESTER") & "," & recsetTCTST3("TS_USER_14") & "," & recsetRun3("RN_RUN_NAME") & "," &  recsetRun3("RN_TESTER_NAME") & "," &  recsetRun3("RN_STATUS") & ","  &  recsetRun3("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun3("RN_USER_07"),",","/") & ","  &  Replace(recsetRun3("RN_USER_06"),",","/") & ","  &  Replace(recsetRun3("RN_USER_05"),",","/") & "," & recsetStep3("ST_STEP_NAME") & "," & recsetStep3("ST_STATUS")
              WriteStuff.WriteLine(Stuff)
              recsetStep3.Next
              Wend
              recsetRun3.Next
              WEND

              '***Run End
              recsetTCTST3.Next
              WEND  'recsetTCTST3.EOR = 0
              Set recsetTCTST3 = NOTHING
              recsetCYCFS1.Next
              WEND 'recsetCYCFS1.EOR = 0

              End If 'recsetFSF("CF_NO_OF_SONS") > 0
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the cycle table to see if it has any test set records
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID, TC_TESTCYCL_ID,TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "

            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0
            comStep4.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun4("RN_RUN_ID")
            Set recsetStep4 = comStep4.Execute
            WHILE recsetStep4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & recsetRun4("RN_RUN_NAME") & "," &  recsetRun4("RN_TESTER_NAME") & "," &  recsetRun4("RN_STATUS") & ","  &  recsetRun4("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun4("RN_USER_07"),",","/") & ","  &  Replace(recsetRun4("RN_USER_06"),",","/") & ","  &  Replace(recsetRun4("RN_USER_05"),",","/") & "," & recsetStep4("ST_STEP_NAME") & "," & recsetStep4("ST_STATUS")
            WriteStuff.WriteLine(Stuff)
            recsetStep4.Next
            Wend
            recsetRun4.Next
            WEND
            Set recsetRun4 = NOTHING
              '***Run End

            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
            Set recsetCYCFS = NOTHING

             '******
           Else 'recsetFS("CF_NO_OF_SONS") > 0 Then
            'Retrieve the Test Set records from CYCLE TABLE
            comCYCFS.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetFS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comTCTST4.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCFS("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
            Set recsetTCTST4 = comTCTST4.Execute
            WHILE recsetTCTST4.EOR = 0
            comRun4.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST4("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST4("TC_TESTCYCL_ID") & " "
            Set recsetRun4 = comRun4.Execute
            WHILE recsetRun4.EOR = 0
            comStep4.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun4("RN_RUN_ID")
            Set recsetStep4 = comStep4.Execute
            WHILE recsetStep4.EOR = 0
            Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & recsetFS("CF_ITEM_NAME") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("CY_CYCLE") & "," & recsetTCTST4("TS_USER_02") & "," & Replace(recsetTCTST4("TS_NAME"),",","") & ","  & recsetTCTST4("TC_TEST_INSTANCE")  & "," & recsetTCTST4("TC_STATUS") & "," & recsetTCTST4("TS_USER_11") & "," & recsetTCTST4("TS_USER_12") & "," & recsetTCTST4("TS_USER_13") & "," & recsetTCTST4("TC_EXEC_DATE") & "," & recsetTCTST4("TC_ACTUAL_TESTER") & "," & recsetTCTST4("TS_USER_14") & "," & recsetRun4("RN_RUN_NAME") & "," &  recsetRun4("RN_TESTER_NAME") & "," &  recsetRun4("RN_STATUS") & ","  &  recsetRun4("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun4("RN_USER_07"),",","/") & ","  &  Replace(recsetRun4("RN_USER_06"),",","/") & ","  &  Replace(recsetRun4("RN_USER_05"),",","/") & "," & recsetStep4("ST_STEP_NAME") & "," & recsetStep4("ST_STATUS")
            WriteStuff.WriteLine(Stuff)
            recsetStep4.Next
            Wend
            recsetRun4.Next
            WEND


            recsetTCTST4.Next
            WEND  'recsetTCTST4.EOR = 0
            Set recsetTCTST4 = NOTHING
            recsetCYCFS.Next
            WEND 'recsetCYCFS.EOR = 0
           End If 'recsetFS("CF_NO_OF_SONS") > 0
           recsetFS.next
           WEND 'recsetFS.EOR = 0
          Else
          'If no child folders exist, retrieve the testset info from CYCLE table
           comCYCF.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetF("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
           Set recsetCYCF = comCYCF.Execute
           WHILE recsetCYCF.EOR = 0
           comTCTST5.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYCF("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
           Set recsetTCTST5 = comTCTST5.Execute
           WHILE recsetTCTST5.EOR = 0
           comRun5.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST5("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST5("TC_TESTCYCL_ID") & " "
           Set recsetRun5 = comRun5.Execute
           WHILE recsetRun5.EOR = 0
           comStep5.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun5("RN_RUN_ID")
           Set recsetStep5 = comStep5.Execute
           WHILE recsetStep5.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME")  & "," & recsetF("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A" & "," & "N/A" & "," & recsetCYCF("CY_CYCLE") & "," & recsetTCTST5("TS_USER_02") & ","  & Replace(recsetTCTST5("TS_NAME"),",","") & ","  & recsetTCTST5("TC_TEST_INSTANCE")  & "," & recsetTCTST5("TC_STATUS") & "," & recsetTCTST5("TS_USER_11") & "," & recsetTCTST5("TS_USER_12") & "," & recsetTCTST5("TS_USER_13") & "," & recsetTCTST5("TC_EXEC_DATE") & "," & recsetTCTST5("TC_ACTUAL_TESTER") & "," & recsetTCTST5("TS_USER_14") & "," & recsetRun5("RN_RUN_NAME") & "," &  recsetRun5("RN_TESTER_NAME") & "," &  recsetRun5("RN_STATUS") & ","  &  recsetRun5("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun5("RN_USER_07"),",","/") & ","  &  Replace(recsetRun5("RN_USER_06"),",","/") & ","  &  Replace(recsetRun5("RN_USER_05"),",","/") & "," & recsetStep5("ST_STEP_NAME") & "," & recsetStep5("ST_STATUS")
           WriteStuff.WriteLine(Stuff)
           recsetStep5.Next
           Wend
           recsetRun5.Next
           WEND

           recsetTCTST5.Next
           WEND  'recsetTCTST5.EOR = 0
           Set recsetTCTST5 = NOTHING
           recsetCYCF.Next
           WEND 'recsetCYCF.EOR
          End If  'recsetF("CF_NO_OF_SONS") > 0
          recsetF.next
          WEND 'recsetF.EOR = 0
         Else
          'If no child folders exist, retrieve the testset info from CYCLE table
          comCYC.CommandText = "Select CY_CYCLE_ID,CY_CYCLE from CYCLE INNER JOIN CYCL_FOLD ON (CY_FOLDER_ID = CF_ITEM_ID) where CY_FOLDER_ID = " & recsetS("CF_ITEM_ID") & " ORDER BY CY_CYCLE"
          Set recsetCYC = comCYC.Execute
          WHILE recsetCYC.EOR = 0
          comTCTST6.CommandText = "SELECT TC_CYCLE_ID,TC_TESTCYCL_ID, TC_TEST_ID,TS_USER_02,TS_NAME,TS_USER_11,TS_USER_12,TS_USER_13,TS_USER_14,TC_STATUS,TC_TESTER_NAME,TC_TEST_INSTANCE,TC_USER_03,TC_PLAN_SCHEDULING_DATE,TC_EXEC_DATE,TC_ACTUAL_TESTER  FROM TESTCYCL,TEST  WHERE tc_cycle_id = " & recsetCYC("CY_CYCLE_ID")  & " and TC_TEST_ID = TS_TEST_ID "
          Set recsetTCTST6 = comTCTST6.Execute
          WHILE recsetTCTST6.EOR = 0
          comRun6.CommandText = "SELECT * FROM RUN WHERE RN_CYCLE_ID = " & recsetTCTST6("TC_CYCLE_ID") & " and RN_TESTCYCL_ID = " & recsetTCTST6("TC_TESTCYCL_ID") & " "
           Set recsetRun6 = comRun6.Execute
           WHILE recsetRun6.EOR = 0
           comStep6.CommandText = "SELECT * FROM STEP INNER JOIN RUN ON (ST_RUN_ID = RN_RUN_ID)where ST_RUN_ID =  " & recsetRun6("RN_RUN_ID")
           Set recsetStep6 = comStep6.Execute
           WHILE recsetStep6.EOR = 0
           Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & recsetS("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "," & recsetCYC("CY_CYCLE") & "," & recsetTCTST6("TS_USER_02") & "," & Replace(recsetTCTST6("TS_NAME"),",","") & ","  & recsetTCTST6("TC_TEST_INSTANCE")  & "," & recsetTCTST6("TC_STATUS") & "," & recsetTCTST6("TS_USER_11") & "," & recsetTCTST6("TS_USER_12") & "," & recsetTCTST6("TS_USER_13") & "," & recsetTCTST6("TC_EXEC_DATE") & "," & recsetTCTST6("TC_ACTUAL_TESTER") & "," & recsetTCTST6("TS_USER_14") & "," & recsetRun6("RN_RUN_NAME") & "," &  recsetRun6("RN_TESTER_NAME") & "," &  recsetRun6("RN_STATUS") & ","  &  recsetRun6("RN_EXECUTION_DATE") & ","  &  Replace(recsetRun6("RN_USER_07"),",","/") & ","  &  Replace(recsetRun6("RN_USER_06"),",","/") & ","  &  Replace(recsetRun6("RN_USER_05"),",","/") & "," & recsetStep6("ST_STEP_NAME") & "," & recsetStep6("ST_STATUS")
           WriteStuff.WriteLine(Stuff)
           recsetStep6.Next
           Wend
           recsetRun6.Next
           WEND


          recsetTCTST6.Next
          WEND  'recsetTCTST6.EOR = 0
          Set recsetTCTST6 = NOTHING
          recsetCYC.Next  'recsetCYC.EOR = 0
          WEND
         End If 'recsetS("CF_NO_OF_SONS") > 0
         recsetS.next
        WEND  'recsetS.EOR = 0

       Else
        'Write to the text file just the Release and the Main Folder
        Stuff = recsetR("CF_ITEM_NAME") & "," & recsetM("CF_ITEM_NAME") & "," & "N/A" & "," & "N/A"  & "," & "N/A" & "," & "N/A" & "N/A"  & "," & "N/A" & "," & "N/A"
        WriteStuff.WriteLine(Stuff)
       End If 'recsetM.EOR

      'Stuff = recsetR("CF_ITEM_NAME") & ", " & recsetM("CF_ITEM_NAME")
      'WriteStuff.WriteLine(Stuff)
      recsetM.next
      WEND  'recsetM.EOR

     End IF ' recsetR("CF_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. c:\ReleaseReport-TestRunStepDetails.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set recsetTCTST6  = nothing
   Set recsetTCTST5 = nothing
   Set recsetTCTST4  = nothing
   Set recsetTCTST3  = nothing
   Set recsetTCTST2 = nothing
   Set recsetTCTST1  = nothing
     Set recsetRun1  = nothing
    Set recsetRun2  = nothing
     Set recsetRun3  = nothing
     Set recsetRun4  = nothing
    Set recsetRun5  = nothing
     Set recsetRun6  = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "RR-TestRunStepDetails"

'**** Release Report TestRunSTEPDetails - End *****************

'Defect Reports
'************************************************
If ActionName = "Defects_By_Release" Then
  If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
 myFSO.DeleteFile "c:\Defects_By_Release.txt"
 Set WriteStuff = myFSO.OpenTextFile("c:\Defects_By_Release.txt", 8, True)
 Stuff = "DefectId" & "," & "Status" & "," & "Release"  & "," & "Environment" & "," & "ITG_Request_Id" & "," & "Detected_Date" & "," & "Detected_By"  & "," & "Assigned_To"  & "," & "Level" & "," & "TestSetName"  & "," & "TestName"  & "," & "TestInstance" & "," & "TestInstanceStatus" & ","  & "RunName" & "," & "RunInstance" & "," & "RunStatus" & "," & "StepName" & "," & "StepStatus"
 WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you the details of all the Defects under each Release")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
  Set tdc   = TDConnection
  Set comR  = tdc.command
  Set comC  = tdc.command
  Set comTC = tdc.command
  Set comRN  = tdc.command
  Set comT  = tdc.command
  Set comS  = tdc.command
  Set comSA  = tdc.command
  'Main Process Logic
  comR.CommandText = "Select BG_BUG_ID,BG_STATUS,BG_RESPONSIBLE,BG_DETECTED_BY,BG_DETECTION_DATE,BG_DETECTED_BY,BG_DETECTION_VERSION,BG_USER_03,BG_REQUEST_ID,LN_ENTITY_TYPE,LN_ENTITY_ID from BUG,LINK where BG_BUG_ID = LN_BUG_ID and BG_Detection_Version = '" & InputBox1 & "'  Order by BG_BUG_ID"
  'MsgBox  comR.CommandText
  Set recsetR = comR.Execute
   WHILE recsetR.EOR = 0
   '   MsgBox recsetR("TCount")
   Select Case recsetR("LN_ENTITY_TYPE")
   Case  "TEST"
        TCount = TCount + 1
         comT.CommandText = "Select TS_NAME from TEST where TS_TEST_ID = " & recsetR("LN_ENTITY_ID") & " "
         Set recsetT = comT.Execute
         WHILE recsetT.EOR = 0
         Stuff = recsetR("BG_BUG_ID") & "," & recsetR("BG_STATUS") & "," & recsetR("BG_DETECTION_VERSION") & "," & recsetR("BG_USER_03") & "," & recsetR("BG_REQUEST_ID") & "," & recsetR("BG_DETECTION_DATE") & "," & recsetR("BG_DETECTED_BY")  & "," & recsetR("BG_RESPONSIBLE") & "," & "TestCase" & "," & "N/A" & "," & Replace(recsetT("TS_NAME"),",","") & "," & "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A"
         WriteStuff.WriteLine(Stuff)
         recsetT.next
         WEND
   Case  "RUN"
         RCount = RCount + 1
         comRN.CommandText = "Select RN_RUN_NAME,RN_STATUS,RN_TEST_INSTANCE,TS_NAME,TC_TEST_INSTANCE,TC_STATUS,CY_CYCLE from RUN,TESTCYCL,TEST,CYCLE where RN_RUN_ID = " & recsetR("LN_ENTITY_ID") & "  and RN_TEST_ID = TS_TEST_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
         Set recsetRN = comRN.Execute
         WHILE recsetRN.EOR = 0
         Stuff = recsetR("BG_BUG_ID") & "," & recsetR("BG_STATUS") & "," & recsetR("BG_DETECTION_VERSION") & "," & recsetR("BG_USER_03") & "," & recsetR("BG_REQUEST_ID")  & "," & recsetR("BG_DETECTION_DATE") & "," & recsetR("BG_DETECTED_BY")  & "," & recsetR("BG_RESPONSIBLE") & "," & recsetR("LN_ENTITY_TYPE") & "," & recsetRN("CY_CYCLE")& "," & Replace(recsetRN("TS_NAME"),",","")& "," & recsetRN("TC_TEST_INSTANCE") &  "," & recsetRN("TC_STATUS")& "," & Replace(recsetRN("RN_RUN_NAME"),",","") & "," & recsetRN("RN_TEST_INSTANCE") & "," & recsetRN("RN_STATUS") & "," &  "N/A" & "," &  "N/A"
         WriteStuff.WriteLine(Stuff)
         recsetRN.next
         WEND
   Case  "STEP"
         SCount = SCount + 1
         comS.CommandText = "Select ST_STEP_NAME,ST_STATUS,RN_RUN_NAME,RN_STATUS,RN_TEST_INSTANCE,TS_NAME,TC_TEST_INSTANCE,TC_STATUS,CY_CYCLE from STEP,RUN,TESTCYCL,TEST,CYCLE where ST_ID = " & recsetR("LN_ENTITY_ID") & "  and ST_RUN_ID = RN_RUN_ID AND ST_TEST_ID = TS_TEST_ID AND RN_CYCLE_ID = CY_CYCLE_ID AND RN_TESTCYCL_ID = TC_TESTCYCL_ID"
         Set recsetS = comS.Execute
         WHILE recsetS.EOR = 0
         Stuff = recsetR("BG_BUG_ID") & "," & recsetR("BG_STATUS") & "," & recsetR("BG_DETECTION_VERSION") & "," & recsetR("BG_USER_03")& "," & recsetR("BG_REQUEST_ID")  & "," & recsetR("BG_DETECTION_DATE") & "," & recsetR("BG_DETECTED_BY")  & "," & recsetR("BG_RESPONSIBLE") & "," & "RunStep" & "," & recsetS("CY_CYCLE")& "," & Replace(recsetS("TS_NAME"),",","")& "," & recsetS("TC_TEST_INSTANCE") &  "," & recsetS("TC_STATUS")& "," & Replace(recsetS("RN_RUN_NAME"),",","") & "," & recsetS("RN_TEST_INSTANCE") & "," & recsetS("RN_STATUS") & "," & recsetS("ST_STEP_NAME")& "," & recsetS("ST_STATUS")
         WriteStuff.WriteLine(Stuff)
         recsetS.next
         WEND

   Case  "BUG"
          BCount = BCount + 1
   Case  "TESTCYCL"
         TCCount = TCCount + 1
         comTC.CommandText = "Select TS_NAME,TC_TEST_INSTANCE,TC_STATUS,CY_CYCLE from TESTCYCL,TEST,CYCLE where TC_TESTCYCL_ID = " & recsetR("LN_ENTITY_ID") & "  and TC_TEST_ID = TS_TEST_ID AND TC_CYCLE_ID = CY_CYCLE_ID"
         Set recsetTC = comTC.Execute
         WHILE recsetTC.EOR = 0
         Stuff = recsetR("BG_BUG_ID") & "," & recsetR("BG_STATUS") & "," & recsetR("BG_DETECTION_VERSION") & "," & recsetR("BG_USER_03") & "," & recsetR("BG_REQUEST_ID") & "," & recsetR("BG_DETECTION_DATE") & "," & recsetR("BG_DETECTED_BY")  & "," & recsetR("BG_RESPONSIBLE") & "," & "TestInstance"  & "," & recsetTC("CY_CYCLE")& "," & Replace(recsetTC("TS_NAME"),",","")& "," & recsetTC("TC_TEST_INSTANCE") & "," & recsetTC("TC_STATUS") & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A"
         WriteStuff.WriteLine(Stuff)
         recsetTC.next
         WEND
   Case  "CYCLE"
         CCount = CCount + 1
         comC.CommandText = "Select CY_CYCLE from CYCLE where CY_CYCLE_ID = " & recsetR("LN_ENTITY_ID") & " "
         Set recsetC = comC.Execute
         WHILE recsetC.EOR = 0
         Stuff = recsetR("BG_BUG_ID") & "," & recsetR("BG_STATUS") & "," & recsetR("BG_DETECTION_VERSION") & "," & recsetR("BG_USER_03") & "," & recsetR("BG_REQUEST_ID")  & "," & recsetR("BG_DETECTION_DATE") & "," & recsetR("BG_DETECTED_BY")  & "," & recsetR("BG_RESPONSIBLE") & "," & "TestSet" & "," & recsetC("CY_CYCLE") & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A"   & "," & "N/A"  & "," & "N/A"
         WriteStuff.WriteLine(Stuff)
         recsetC.next
         WEND

   Case Else
         ElseCount = ElseCount + 1
   End Select
   recsetR.next
   Wend

   ' Select the standalone defects that are not tied to the Test Entities
   comSA.CommandText = "Select BG_BUG_ID,BG_STATUS,BG_RESPONSIBLE,BG_DETECTED_BY,BG_DETECTION_DATE,BG_DETECTED_BY,BG_DETECTION_VERSION,BG_USER_03,BG_REQUEST_ID FROM BUG where BG_BUG_ID NOT IN (SELECT LN_BUG_ID FROM LINK) and BG_Detection_Version = '" & InputBox1 & "'  Order by BG_BUG_ID"
   Set recsetSA = comSA.Execute
   WHILE recsetSA.EOR = 0
   Stuff = recsetSA("BG_BUG_ID") & "," & recsetSA("BG_STATUS") & "," & recsetSA("BG_DETECTION_VERSION") & "," & recsetSA("BG_USER_03") & "," & recsetSA("BG_REQUEST_ID")  & "," & recsetSA("BG_DETECTION_DATE") & "," & recsetSA("BG_DETECTED_BY")  & "," & recsetSA("BG_RESPONSIBLE") & "," & "N/A" & "," & "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A" & "," &  "N/A"   & "," & "N/A"  & "," & "N/A"
   WriteStuff.WriteLine(Stuff)
   recsetSA.next
   Wend

   'Clear the recordsets
   Set recsetR = nothing
   Set recsetC = nothing
   Set recsetTC = nothing
   Set recsetRN = nothing
   Set recsetT = nothing
   Set recsetS = nothing
   Set recsetSA = nothing
  End If
   'MsgBox "TEST:" & TCount & " RUN:" & RCount & " STEP:" & SCount & " BUG:" & BCount & " TestCycle:" & TCCount & " Cycle:" & CCount & " Others:" & ElseCount
   WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
    MsgBox "Report dowload Complete. C:\Defects_By_Release.txt"
End If
End If 'Action- Defects_By_Release


'***TEST PLAN REPORTS *****************************

'**** Release Report No Planned Info - Begin ***************
If ActionName = "Act_NoPlanned" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\ReleaseReport-NoPlanned.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-NoPlanned.txt", 8, True)
  Stuff = "Test Cases with No Planned Info"
  WriteStuff.WriteLine(Stuff)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Case" & ", " & "Designer" & ", " & "Scripter"  & ", " & "Planned Tester" & ", " & "Planned Start Date"  & ", " & "Planned End Date"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give list of Scenarios with no Planned Info under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\No Planned Info Report")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = 2 and AL_DESCRIPTION = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     'If recsetR("AL_NO_OF_SONS") > 0 Then

      InputBox2 = InputBox ("Enter the name of the SubFolder. Leave the field blank if you want to report on all the sub folders","Release\No Planned Info Report")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      ELse
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = '" & InputBox2 & "' and AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      End If

      Set recsetM = comM.Execute

      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       'If recsetM("AL_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = 'HHSC' and AL_FATHER_ID = " & recsetM("AL_ITEM_ID")
        Set recsetS = comS.Execute

        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         'If recsetS("AL_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetS("AL_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          'If recsetF("AL_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetF("AL_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           'If recsetFS("AL_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFS("AL_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              'If recsetFSF("AL_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFSF("AL_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the TEST table if any
              comCYCFS2.CommandText = "Select al_item_id,TS_TEST_ID,TS_NAME,TS_RESPONSIBLE,TS_USER_14,TS_USER_11,TS_USER_12,TS_USER_13 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where ((TS_USER_11 IS NULL) OR (TS_USER_12 IS NULL) OR (TS_USER_13 IS NULL)) AND TS_SUBJECT = " & recsetFSF1("AL_ITEM_ID") & ""
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", " & recsetFSF("AL_DESCRIPTION") & ", " & recsetFSF1("AL_DESCRIPTION") &  ", " & Replace(recsetCYCFS2("TS_NAME"),",","") &  ", " &  recsetCYCFS2("TS_RESPONSIBLE") &  ", " &  recsetCYCFS2("TS_USER_14") &  ", " & recsetCYCFS2("TS_USER_11") &  ", " & recsetCYCFS2("TS_USER_12")  &  ", " & recsetCYCFS2("TS_USER_13")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select TS_NAME,TS_RESPONSIBLE,TS_USER_14,TS_USER_11,TS_USER_12,TS_USER_13 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where ((TS_USER_11 IS NULL) OR (TS_USER_12 IS NULL) OR (TS_USER_13 IS NULL)) AND TS_SUBJECT = " & recsetFSF("AL_ITEM_ID") & " "
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", " & recsetFSF("AL_DESCRIPTION")& ", "  & "N/A"  &  ", " & Replace(recsetCYCFS1("TS_NAME"),",","") &  ", " & recsetCYCFS1("TS_RESPONSIBLE")&  ", " & recsetCYCFS1("TS_USER_14")&  ", " & recsetCYCFS1("TS_USER_11")  &  ", " & recsetCYCFS1("TS_USER_12")&  ", " & recsetCYCFS1("TS_USER_13")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND
              Set recsetCYCFS1 = Nothing
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the TEST table to see if it has any test records
            'Retrieve the Test records from TEST TABLE
            comCYCFS.CommandText = "Select TS_NAME,TS_RESPONSIBLE,TS_USER_14,TS_USER_11,TS_USER_12,TS_USER_13 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where ((TS_USER_11 IS NULL) OR (TS_USER_12 IS NULL) OR (TS_USER_13 IS NULL)) AND TS_SUBJECT = " & recsetFS("AL_ITEM_ID") & ""
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", "  & "N/A"  & ", "  & "N/A" & ", " & Replace(recsetCYCFS("TS_NAME"),",","") & ", " & recsetCYCFS("TS_RESPONSIBLE")& ", " & recsetCYCFS("TS_USER_14") & ", " & recsetCYCFS("TS_USER_11") & ", " & recsetCYCFS("TS_USER_12")& ", " & recsetCYCFS("TS_USER_13")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND
            Set recsetCYCFS = NOTHING
          recsetFS.next
           WEND 'recsetFS.EOR = 0
          recsetF.next
          WEND 'recsetF.EOR = 0
          recsetS.next
        WEND  'recsetS.EOR = 0
       recsetM.next
      WEND  'recsetM.EOR

     'End IF ' recsetR("AL_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-NoPlanned.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "Act-NoPlanned"
'**** Release Report No Planned - End ***************


'***** Test Scenarios Count *************
'**** Release Report TestScenariosCount - Begin ***************
If ActionName = "TestScenariosCount" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\ReleaseReport-TestScenariosCount.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestScenariosCount.txt", 8, True)
  Stuff = "Test Scenarios Count"
  WriteStuff.WriteLine(Stuff)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Designer"  & ", " & "Planned Tester" & ", " & "Scripter"  & ", " & "Count"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you a count of Test Scenarios for each Tester under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\Test Scenarios Count")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic


   comR.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = 2 and AL_DESCRIPTION = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     'If recsetR("AL_NO_OF_SONS") > 0 Then

      InputBox2 = InputBox ("Enter the name of the SubFolder. Leave the field blank if you want to report on all the sub folders","Release\Test Scenarios Count")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      ELse
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = '" & InputBox2 & "' and AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      End If

      Set recsetM = comM.Execute

      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       'If recsetM("AL_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = 'HHSC' and AL_FATHER_ID = " & recsetM("AL_ITEM_ID")
        Set recsetS = comS.Execute

        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         'If recsetS("AL_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetS("AL_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          'If recsetF("AL_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetF("AL_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           'If recsetFS("AL_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFS("AL_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              'If recsetFSF("AL_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFSF("AL_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the TEST table if any
              comCYCFS2.CommandText = "Select TS_RESPONSIBLE,TS_USER_11,TS_USER_14, Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF1("AL_ITEM_ID") & "  Group By TS_RESPONSIBLE,TS_USER_11,TS_USER_14"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", " & recsetFSF("AL_DESCRIPTION") & ", " & recsetFSF1("AL_DESCRIPTION") &  ", " &  recsetCYCFS2("TS_RESPONSIBLE")&  ", " & recsetCYCFS2("TS_USER_11") &  ", " & recsetCYCFS2("TS_USER_14") &  ", " &  recsetCYCFS2("TestCount")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select TS_RESPONSIBLE,TS_USER_11,TS_USER_14,Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF("AL_ITEM_ID") & " Group By TS_RESPONSIBLE,TS_USER_11,TS_USER_14"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", " & recsetFSF("AL_DESCRIPTION")& ", "  & "N/A"  &  ", " & recsetCYCFS1("TS_RESPONSIBLE")&  ", " & recsetCYCFS1("TS_USER_11") &  ", " & recsetCYCFS1("TS_USER_14") &  ", " & recsetCYCFS1("TestCount")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND
              Set recsetCYCFS1 = Nothing
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the TEST table to see if it has any test records
            'Retrieve the Test records from TEST TABLE
            comCYCFS.CommandText = "Select TS_RESPONSIBLE,TS_USER_11,TS_USER_14, Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFS("AL_ITEM_ID") & " GROUP BY TS_RESPONSIBLE ,TS_USER_11,TS_USER_14"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("TS_RESPONSIBLE")&  ", " & recsetCYCFS("TS_USER_11") &  ", " & recsetCYCFS("TS_USER_14") & ", " & recsetCYCFS("TestCount")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND
            Set recsetCYCFS = NOTHING
          recsetFS.next
           WEND 'recsetFS.EOR = 0
          recsetF.next
          WEND 'recsetF.EOR = 0
          recsetS.next
        WEND  'recsetS.EOR = 0
       recsetM.next
      WEND  'recsetM.EOR

     'End IF ' recsetR("AL_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-TestScenariosCount.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "TestScenariosCount"
'**** Release Report TestScenariosCount - End ***************


'***** Test Scenarios Steps Count *************
'**** Release Report TestScenariosStepsCount - Begin ***************
If ActionName = "TestScenariosStepsCount" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\ReleaseReport-TestScenariosStepsCount.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-TestScenariosStepsCount.txt", 8, True)
  Stuff = "Test Scenarios Steps Count"
  WriteStuff.WriteLine(Stuff)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Designer"  & ", " & "Test Scenario"  & ", " & "Count"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you a count of Test Scenarios Steps for each Tester under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\Test Scenarios Steps Count")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command
   Set comCYCDS1  = tdc.command
   Set comCYCDS2 = tdc.command
    Set comCYCDS3 = tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = 2 and AL_DESCRIPTION = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     'If recsetR("AL_NO_OF_SONS") > 0 Then

      InputBox2 = InputBox ("Enter the name of the SubFolder. Leave the field blank if you want to report on all the sub folders","Release\Test Scenarios Steps Count")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      ELse
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = '" & InputBox2 & "' and AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      End If

      Set recsetM = comM.Execute

      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       'If recsetM("AL_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = 'HHSC' and AL_FATHER_ID = " & recsetM("AL_ITEM_ID")
        Set recsetS = comS.Execute

        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         'If recsetS("AL_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetS("AL_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          'If recsetF("AL_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetF("AL_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           'If recsetFS("AL_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFS("AL_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              'If recsetFSF("AL_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFSF("AL_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the TEST table if any
              comCYCFS2.CommandText = "Select TS_TEST_ID,TS_RESPONSIBLE,TS_NAME from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF1("AL_ITEM_ID") & " "
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
                comCYCDS1.CommandText = "Select Count(*) as DSCount from DESSTEPS INNER JOIN TEST ON (TS_TEST_ID = DS_TEST_ID) WHERE DS_TEST_ID = " & recsetCYCFS2("TS_TEST_ID")
                'msgbox  comCYCDS1.CommandText

                Set recsetCYCDS1 = comCYCDS1.Execute
                If  recsetCYCDS1.EOR = 0 Then

                Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & "," & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & "," & recsetFSF("AL_DESCRIPTION") & "," & recsetFSF1("AL_DESCRIPTION") &  "," &  recsetCYCFS2("TS_RESPONSIBLE") &  "," &  Replace(recsetCYCFS2("TS_NAME"),",","") &  "," &  recsetCYCDS1("DSCount")
                WriteStuff.WriteLine(Stuff)
                End If

                recsetCYCFS2.Next
              WEND
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select TS_TEST_ID,TS_RESPONSIBLE,TS_NAME from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF("AL_ITEM_ID") & ""
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              comCYCDS2.CommandText = "Select Count(*) as DSCount from DESSTEPS INNER JOIN TEST ON (TS_TEST_ID = DS_TEST_ID) WHERE DS_TEST_ID = " & recsetCYCFS1("TS_TEST_ID")
                'msgbox  comCYCDS1.CommandText

              Set recsetCYCDS2 = comCYCDS2.Execute
              If  recsetCYCDS2.EOR = 0 Then
              Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", " & recsetFSF("AL_DESCRIPTION")& ", "  & "N/A"  &  ", " & recsetCYCFS1("TS_RESPONSIBLE")&  ", " & Replace(recsetCYCFS1("TS_Name"),",","") &  ", " & recsetCYCDS2("DSCount")
              WriteStuff.WriteLine(Stuff)
              End If
              recsetCYCFS1.Next
              WEND
              Set recsetCYCFS1 = Nothing
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the TEST table to see if it has any test records
            'Retrieve the Test records from TEST TABLE
            comCYCFS.CommandText = "Select TS_TEST_ID,TS_RESPONSIBLE,TS_NAME from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFS("AL_ITEM_ID") & ""
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            comCYCDS3.CommandText = "Select Count(*) as DSCount from DESSTEPS INNER JOIN TEST ON (TS_TEST_ID = DS_TEST_ID) WHERE DS_TEST_ID = " & recsetCYCFS("TS_TEST_ID")
            Set recsetCYCDS3 = comCYCDS3.Execute
            If  recsetCYCDS3.EOR = 0 Then
            Stuff = recsetR("AL_DESCRIPTION") & ", " & recsetM("AL_DESCRIPTION") & ", " & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & ", " & recsetFS("AL_DESCRIPTION") & ", "  & "N/A"  & ", "  & "N/A" & ", " & recsetCYCFS("TS_RESPONSIBLE")& ", " & Replace(recsetCYCFS("TS_NAME"),",","") &  ", " & recsetCYCDS3("DSCount")
            WriteStuff.WriteLine(Stuff)
            End If
            recsetCYCFS.Next
            WEND
            Set recsetCYCFS = NOTHING
          recsetFS.next
           WEND 'recsetFS.EOR = 0
          recsetF.next
          WEND 'recsetF.EOR = 0
          recsetS.next
        WEND  'recsetS.EOR = 0
       recsetM.next
      WEND  'recsetM.EOR

     'End IF ' recsetR("AL_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-TestScenariosStepsCount.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "TestScenariosCount"
'**** Release Report TestScenariosCount - End ***************

'**** Release Report Planned Date ScenariosCount - Begin ***************
If ActionName = "Act_PlannedDtCount" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\ReleaseReport-PlannedDtScenariosCount.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-PlannedDtScenariosCount.txt", 8, True)
  Stuff = "Planned Date Test Scenarios Count"
  WriteStuff.WriteLine(Stuff)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Planned Start Date"  & ", " & "Designer"  & ", " & "Planned Tester" & ", " & "Scripter"  & ", " & "Count"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will give you a count of Planned Date Test Scenarios for each Tester under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\Planned Date Test Scenarios Count")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = 2 and AL_DESCRIPTION = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     'If recsetR("AL_NO_OF_SONS") > 0 Then

      InputBox2 = InputBox ("Enter the name of the SubFolder. Leave the field blank if you want to report on all the sub folders","Release\Planned Date Scenarios Count")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      ELse
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = '" & InputBox2 & "' and AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      End If
     Set recsetM = comM.Execute

      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       'If recsetM("AL_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = 'HHSC' and AL_FATHER_ID = " & recsetM("AL_ITEM_ID")
        Set recsetS = comS.Execute

        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         'If recsetS("AL_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetS("AL_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          'If recsetF("AL_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetF("AL_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           'If recsetFS("AL_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFS("AL_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              'If recsetFSF("AL_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFSF("AL_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the TEST table if any
              comCYCFS2.CommandText = "Select TS_USER_12,TS_RESPONSIBLE,TS_USER_11,TS_USER_14, Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF1("AL_ITEM_ID") & "  Group By TS_USER_12,TS_RESPONSIBLE,TS_USER_11,TS_USER_14"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & "," & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & "," & recsetFSF("AL_DESCRIPTION") & "," & recsetFSF1("AL_DESCRIPTION") &  "," &  recsetCYCFS2("TS_USER_12") &  "," &  recsetCYCFS2("TS_RESPONSIBLE")&  "," & recsetCYCFS2("TS_USER_11") &  "," & recsetCYCFS2("TS_USER_14") &  "," &  recsetCYCFS2("TestCount")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select TS_USER_12,TS_RESPONSIBLE,TS_USER_11,TS_USER_14,Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF("AL_ITEM_ID") & " Group By TS_USER_12,TS_RESPONSIBLE,TS_USER_11,TS_USER_14"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & "," & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & "," & recsetFSF("AL_DESCRIPTION")& ","  & "N/A"  &  "," & recsetCYCFS1("TS_USER_12") &  "," & recsetCYCFS1("TS_RESPONSIBLE")&  "," & recsetCYCFS1("TS_USER_11") &  "," & recsetCYCFS1("TS_USER_14") &  "," & recsetCYCFS1("TestCount")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND
              Set recsetCYCFS1 = Nothing
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the TEST table to see if it has any test records
            'Retrieve the Test records from TEST TABLE
            comCYCFS.CommandText = "Select TS_USER_12,TS_RESPONSIBLE,TS_USER_11,TS_USER_14, Count(*) as TestCount from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFS("AL_ITEM_ID") & " GROUP BY TS_USER_12,TS_RESPONSIBLE ,TS_USER_11,TS_USER_14"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & ","  & "N/A"  & ","  & "N/A" & "," & recsetCYCFS("TS_USER_12")&  "," & recsetCYCFS("TS_RESPONSIBLE")&  "," & recsetCYCFS("TS_USER_11") &  "," & recsetCYCFS("TS_USER_14") & "," & recsetCYCFS("TestCount")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND
            Set recsetCYCFS = NOTHING
          recsetFS.next
           WEND 'recsetFS.EOR = 0
          recsetF.next
          WEND 'recsetF.EOR = 0
          recsetS.next
        WEND  'recsetS.EOR = 0
       recsetM.next
      WEND  'recsetM.EOR

     'End IF ' recsetR("AL_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-PlannedDtScenariosCount.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "TestScenariosCount"
'**** Release Report TestScenariosCount - End ***************
'*************************************************


'**** Release Report Planned Date ScenariosCount - Begin ***************
If ActionName = "ScenariosByPlannedDt" Then
 If User.IsInGroup("Custom Reports")Then
 Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\ReleaseReport-ScenariosByPlannedDt.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\ReleaseReport-ScenariosByPlannedDt.txt", 8, True)
  Stuff = "Scenarios By Planned Date"
  WriteStuff.WriteLine(Stuff)
  Stuff = "Level 1" & ", " & "Level 2" & ", " & "Level 3"  & ", " & "Level 4" & ", " & "Level 5" & ", " & "Level 6"  & ", " & "Level 7"  & ", " & "Test Scenario" & ", " & "Planned Start Date" & ", " & "Planned End Date"  & ", " & "Designer"  & ", " & "Planned Tester" & ", " & "Scripter"
  WriteStuff.WriteLine(Stuff)

  InputBox1 = InputBox ("The report will list all Scenarios by Planned Date for each Tester under each Release Folder by Release Number & HHSC Sub folders. Enter the Release Number","Release\Scenarios by Planned Date")

  If len(InputBox1) > 0 Then
   'Cycl_Fold table command sets
   Set tdc = TDConnection
   Set comR = tdc.command
   Set comM = tdc.command
   Set comS = tdc.command
   Set comF = tdc.command
   Set comFS = tdc.command
   Set comFSF =tdc.command
   Set comFSF1 =tdc.command
   Set comFSF2 =tdc.command

   'Cycle table command sets
   Set comCYCFS2 = tdc.command
   Set comCYCFS1 = tdc.command
   Set comCYCFS = tdc.command
   Set comCYCF = tdc.command
   Set comCYC = tdc.command
   'Main Process Logic
   comR.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = 2 and AL_DESCRIPTION = '" & InputBox1 & "' "
   Set recsetR = comR.Execute

   If recsetR.EOR = 0 Then
    'MsgBox "Folder Does  exist"
    'Check to see if the parent folder has any child folders
     'If recsetR("AL_NO_OF_SONS") > 0 Then

      'Check to see if the user wants to run the report for one sub folder or all

      InputBox2 = InputBox ("Enter the name of the SubFolder. Leave the field blank if you want to report on all the sub folders","Release\Scenarios by Planned Date")
      If len(InputBox2) <= 0 Then
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      ELse
      comM.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = '" & InputBox2 & "' and AL_FATHER_ID = '" & recsetR("AL_ITEM_ID") & "' Order by AL_DESCRIPTION"
      End If

      Set recsetM = comM.Execute

      WHILE recsetM.EOR = 0
       'Check if the Main Parent folder has any child folders
       'If recsetM("AL_NO_OF_SONS") > 0 Then
        comS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_DESCRIPTION = 'HHSC' and AL_FATHER_ID = " & recsetM("AL_ITEM_ID")
        Set recsetS = comS.Execute

        WHILE recsetS.EOR = 0
         'Check if the Subfolder HHSC has any child folders
         'If recsetS("AL_NO_OF_SONS") > 0 Then
          comF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetS("AL_ITEM_ID")
          Set recsetF = comF.Execute
          WHILE recsetF.EOR = 0
          'Check if the Folder has any child folders
          'If recsetF("AL_NO_OF_SONS") > 0 Then
           comFS.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetF("AL_ITEM_ID")
           Set recsetFS = comFS.Execute
           WHILE recsetFS.EOR = 0

           'Check if the Folder/Sub has any child folders
           'If recsetFS("AL_NO_OF_SONS") > 0 Then
            'Check The folder Table to retrieve the child folders
            comFSF.CommandText = "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFS("AL_ITEM_ID")
            Set recsetFSF = comFSF.Execute
            WHILE recsetFSF.EOR = 0
              'If recsetFSF("AL_NO_OF_SONS") > 0 Then
              comFSF1.CommandText =  "Select AL_ITEM_ID,AL_FATHER_ID,AL_DESCRIPTION,AL_NO_OF_SONS from ALL_LISTS where AL_FATHER_ID = " & recsetFSF("AL_ITEM_ID")
              Set recsetFSF1 = comFSF1.Execute
              WHILE recsetFSF1.EOR = 0
              'LAST record to retrieve folder level 7
              'Retrieve records from the TEST table if any
              comCYCFS2.CommandText = "Select TS_NAME,TS_USER_12,TS_USER_13,TS_RESPONSIBLE,TS_USER_11,TS_USER_14 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF1("AL_ITEM_ID") & "  Order by TS_USER_11"
              Set recsetCYCFS2 = comCYCFS2.Execute
              WHILE recsetCYCFS2.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & "," & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & "," & recsetFSF("AL_DESCRIPTION") & "," & recsetFSF1("AL_DESCRIPTION") &  "," &  Replace(recsetCYCFS2("TS_NAME"),",","")  &  "," &  recsetCYCFS2("TS_USER_12") &  "," &  recsetCYCFS2("TS_USER_13") &  "," &  recsetCYCFS2("TS_RESPONSIBLE")&  "," & recsetCYCFS2("TS_USER_11") &  "," & recsetCYCFS2("TS_USER_14")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS2.Next
              WEND
              recsetFSF1.Next
              WEND 'recsetFSF1.EOR = 0

              '****
              'Retrieve the Test Set records from CYCLE TABLE
              comCYCFS1.CommandText = "Select TS_NAME,TS_USER_12,TS_USER_13,TS_RESPONSIBLE,TS_USER_11,TS_USER_14 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFSF("AL_ITEM_ID") & " Order by TS_USER_11"
              Set recsetCYCFS1 = comCYCFS1.Execute
              WHILE recsetCYCFS1.EOR = 0
              Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & "," & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & "," & recsetFSF("AL_DESCRIPTION")& ","  & "N/A"  &  "," &  Replace(recsetCYCFS1("TS_NAME"),",","") &  "," & recsetCYCFS1("TS_USER_12") &  "," & recsetCYCFS1("TS_USER_13") &  "," & recsetCYCFS1("TS_RESPONSIBLE")&  "," & recsetCYCFS1("TS_USER_11") &  "," & recsetCYCFS1("TS_USER_14")
              WriteStuff.WriteLine(Stuff)
              recsetCYCFS1.Next
              WEND
              Set recsetCYCFS1 = Nothing
              recsetFSF.next
            WEND  ' recsetFSF.EOR = 0
            '*******
            'Check the TEST table to see if it has any test records
            'Retrieve the Test records from TEST TABLE
            comCYCFS.CommandText = "Select TS_NAME,TS_USER_12,TS_USER_13,TS_RESPONSIBLE,TS_USER_11,TS_USER_14 from TEST INNER JOIN ALL_LISTS ON (TS_SUBJECT = AL_ITEM_ID) where TS_SUBJECT = " & recsetFS("AL_ITEM_ID") & " Order by TS_USER_11"
            Set recsetCYCFS = comCYCFS.Execute
            WHILE recsetCYCFS.EOR = 0
            Stuff = recsetR("AL_DESCRIPTION") & "," & recsetM("AL_DESCRIPTION") & "," & recsetS("AL_DESCRIPTION")  & ", " & recsetF("AL_DESCRIPTION") & "," & recsetFS("AL_DESCRIPTION") & ","  & "N/A"  & ","  & "N/A" &  "," &  Replace(recsetCYCFS("TS_NAME"),",","") & "," & recsetCYCFS("TS_USER_12") & "," & recsetCYCFS("TS_USER_13") &  "," & recsetCYCFS("TS_RESPONSIBLE")&  "," & recsetCYCFS("TS_USER_11") &  "," & recsetCYCFS("TS_USER_14")
            WriteStuff.WriteLine(Stuff)
            recsetCYCFS.Next
            WEND
            Set recsetCYCFS = NOTHING
          recsetFS.next
           WEND 'recsetFS.EOR = 0
          recsetF.next
          WEND 'recsetF.EOR = 0
          recsetS.next
        WEND  'recsetS.EOR = 0
       recsetM.next
      WEND  'recsetM.EOR

     'End IF ' recsetR("AL_NO_OF_SONS") > 0
      MsgBox "Report dowload Complete. C:\ReleaseReport-ScenariosByPlannedDt.txt"
   Else
    MsgBox "Release Folder Does Not exist"
   End If  'recsetR.EOR

   Set recsetR = nothing
   Set recsetM = nothing
   Set recsetS = nothing
   Set recsetF = nothing
   Set recsetFS = nothing
   Set recsetFSF = nothing
   Set recsetFSF1 = nothing
   Set  recsetCYCFS2 = nothing
   Set  recsetCYCFS1 = nothing
   Set  recsetCYCFS = nothing
   Set  recsetCYCF = nothing
   Set  recsetCYC = nothing
   Set tdc = nothing
  End If 'len(InputBox1) > 0
  WriteStuff.Close
  SET WriteStuff = NOTHING
  SET myFSO = NOTHING
 End If 'User.IsInGroup("Custom Reports")
End If  'ActionName = "ScenariosByPlannedDt"
'**** Release Report ScenariosByPlannedDt - End ***************

If ActionName = "Act_UATUsers" Then
 If User.IsInGroup("HHSC-UAT") or User.IsInGroup("Superuser") Then
  Dim custuser
  Set myFSO = CreateObject("Scripting.FileSystemObject")
  myFSO.DeleteFile "c:\HHSC_UAT_UsersList.txt"
  Set WriteStuff = myFSO.OpenTextFile("c:\HHSC_UAT_UsersList.txt", 8, True)
  Set tdc = TDConnection
  Set cust = tdc.Customization
        cust.Load
        Set custUsers = cust.Users
        Set grps = cust.UsersGroups
        Set groups = grps.group("HHSC-UAT")
        set custlists = cust.lists
        set custlist = groups.UsersList
        For i= 1 to custlist.count
          set custuser = custlist.item(i)
           WriteStuff.WriteLine(custuser.Name)
          'myFSO.WriteLine custuser.Name
        Next
        'added below code based on ER156085...twenger
        set groups = grps.group("HHSC-Runs")
        set custlists = cust.lists
        set custlist = groups.UsersList
        for i= 1 to custlist.count
          set custuser = custlist.item(i)
           WriteStuff.WriteLine(custuser.Name)
        next
        'added below code based on ER156085...twenger
        set groups = grps.group("Superuser")
        set custlists = cust.lists
        set custlist = groups.UsersList
        for i= 1 to custlist.count
          set custuser = custlist.item(i)
           WriteStuff.WriteLine(custuser.Name)
        next

  Set tdc = nothing

  WriteStuff.Close
  Set WriteStuff = NOTHING
  Set myFSO = NOTHING
  MsgBox "Report dowload Complete. C:\HHSC_UAT_UsersList.txt"
 End If 'User.IsInGroup("Custom Reports")
End If 'Act_UATUsers


ActionCanExecute = DefaultRes
On Error GoTo 0
End Function

Function CanLogout
  On Error Resume Next

  CanLogout = DefaultRes
  On Error GoTo 0
End Function

Sub EnterModule
  'Use ActiveModule and ActiveDialogName to get
  'the current context.
  On Error Resume Next
  'MsgBox ActiveModule + "  ActiveDialogName   " + ActiveDialogName
  if ActiveModule = "Defects" and not User.IsInGroup("TDAdmin") then
     Actions.Action("acnCopy").Enabled = False
     Actions.Action("acnCopy").Visible = False
  end if
  On Error GoTo 0
End Sub

Function CanLogin(DomainName, ProjectName, UserName)
  On Error Resume Next
  'if ProjectName = "Integrated_Eligibility" then
   ' if UserName <> "tim.wenger" then
   '   msgbox "Currently " & ProjectName & " is NOT available for use. Will notify team leads when you are allowed to login."
   '   CanLogin = False
   ' else
   '   CanLogin = True
   ' end if
  'end if
  CanLogin = DefaultRes
  On Error GoTo 0
End Function
