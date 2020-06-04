AddReference "QuickstartiLogicLibrary.dll"
'DISCLAIMER:
'---------------------------------
'In any case, code, templates, and snippets of this solution are of "work in progress" character.
'Neither Markus Koechl, nor Autodesk represents that these samples are reliable, accurate, complete, or otherwise valid. 
'Accordingly, those configuration samples are provided “as is” with no warranty of any kind and you use the applications at your own risk.
Class TestShared
	
	Sub main
		
		Dim partpath As String
		Dim ibstring As String
		ibstring = InputBox("Enter iam or ipt name", "Quick place")
		If Not ibstring Is Nothing Then
			If ibstring.Count > 4 Then
				If ibstring.Substring(ibstring.Length - 4) = ".iam" Or ibstring.Substring(ibstring.Length - 4) = ".ipt" Then
					partpath = getVaultFile(ibstring)
				Else
					Logger.Error("Incorrect input")
					Exit Sub
				End If
			End If
		End If
	
		If Not partpath Is Nothing Then ''Has been found and is downloaded
			Logger.Info("Part loaded..." )
			Dim oAsmCompDef As AssemblyComponentDefinition
			oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition
			Dim oTG As TransientGeometry = ThisApplication.TransientGeometry
			Dim oMatrix As Matrix = oTG.CreateMatrix
			oMatrix.SetTranslation(oTG.CreateVector(0, 0, 0))
			oOccurrence = oAsmCompDef.Occurrences.Add(partpath, oMatrix)
			oOccurrence.Grounded = True
		Else
			Logger.Error("Could not find file")				
		End If
			
	End Sub
	
	Function getVaultFile(filename As String) As String
		
		Dim mSearchParams As New System.Collections.Generic.Dictionary(Of String, String) 'declare params
		Dim iLogicVault As New QuickstartiLogicLibrary.QuickstartiLogicLib
		mSearchParams.Add("File Name", filename)	'applies to file 001002.ipt

		'enable iLogicVault commands and validate user's login state
		
		If iLogicVault.LoggedIn = False
			Logger.Error("Not Logged In to Vault! - Login first and repeat executing this rule.")
			Return Nothing
		End If
		break
		'returns full file name in local working folder (download enforces override, if local file exists)
		Dim mVaultFile As String = iLogicVault.GetFileBySearchCriteria(mSearchParams)
		If mVaultFile Is Nothing Then
			Return Nothing
		Else
			Return mVaultFile
		End If
	End Function
End Class



	
