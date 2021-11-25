Attribute VB_Name = "Module_GetText"
' ==================================================================
' == XL TOOLBOX   (c) 2008-2013  Daniel Kraus   Licensed under GPLv2
' ==================================================================
' == DanielsXLToolbox.Module_GetText
' ==
' == Created: 04-Nov-11 21:08
' ==
' == ===============================================================
' == See: http://vbagettext.sf.net
' == ===============================================================
' ==
' == This module implements GNU GetText-like behavior to permit
' == very basic I18N and L10N of Excel VBA projects.
' == "Install" the module by importing it into the VBA project that
' == you wish to equip with I18N.
' ==
' == In modules/classes/forms, when text is needed, use t(text)
' == to translate.
' ==
' == To get a list of available translations, use GetLanguages.
' ==
' == To set a translation, use SetLanguage.
' ==
' == To automatically set the translation according to Excel's
' == UI language, use DetectLanguage.
' ==
' == To retrieve strings from the current VBA project, use the
' == CollectStrings procedure. This will output a 'po' file
' == in a 'locales' directory as the current project. (The name
' == of the subdirectory can be changed by editing the LOCALE_DIR
' == constant).

Option Private Module
Option Explicit

    ' The default languge of this addin.
    Public Const DEFAULT_LANGUAGE = "en-US"
    
    Type tLocaleInfo
        Name As String
        Code As String * 2
        LocaleCode As String * 5
        id As Long
    End Type
    
    ' Adjust the next two constants for use in your project
    Private Const LOCALE_DIR = "resource"             ' where to locate the translation files
    Private Const LOCALE_BASE_NAME = "inoRound" ' base name of the PO files
    
    Private Const MENU_CONTEXT = "Menu"
    Private Const EOT = 4
    
    Private mCurrentLang As String
    Private mNumLocales As Long
    Private mLocales() As tLocaleInfo
    Private mPoFile As clsGetTextFile


#If VBA7 Then
    Private Declare PtrSafe Function compute_hashval Lib "xltoolbox.dll" (ByVal text As String, ByVal keylen As Long) As Long
#Else
    Private Declare Function compute_hashval Lib "xltoolbox.dll" (ByVal text As String, ByVal keylen As Long) As Long
#End If



Function RegisterXLToolboxDLL() As Boolean
' This function will register the path to the xltoolbox.dll
' file which is required for some computations. Call this
' function once before using the vbagettext functions.
' The DLL is assumed primarily in the current workbook's folder,
' then in a subfolder called 'resource'.
    On Error Resume Next
    Dim strPath As String
    strPath = AddPathSep(ThisWorkbook.path)
    If Not FileExists(strPath & "xltoolbox.dll") Then
        strPath = AddPathSep(strPath & "resource")
        If Not FileExists(strPath & "xltoolbox.dll") Then Exit Function
    End If
    AddDLLPath strPath
    If Err.LastDllError = 0 Then
        RegisterXLToolboxDLL = True
    Else
        If Err.LastDllError Then Debug.Print "DLL error #" & Err.LastDllError & _
            "occurred in vbagettext::RegisterXLToolboxDLL()"
    End If
End Function

Function UnregisterXLToolboxDLL() As Boolean
    On Error Resume Next
    UnregisterXLToolboxDLL = RemoveDLLPath
    If Err.LastDllError = 0 Then
        UnregisterXLToolboxDLL = True
    Else
        If Err.LastDllError Then Debug.Print "DLL error #" & Err.LastDllError & _
            "occurred in vbagettext::UnregisterXLToolboxDLL()"
    End If
End Function


Function t(ByVal text As String, ParamArray params() As Variant) As String
' The main GetText function: 't' as in 'translate'.
' (VBA does not allow function names that consist of a single underscore.)
' The function takes 'text' and attempts to look it up against a translation
' file. Returns the translated version of 'text' if one was found.
' Any occurrence of '{}' in the string will be replace with a parameter
' handed over in params().

    Dim s_ As String
    Dim i As Long
    Dim j As Long
    Dim lb As Long
    Dim ub As Long
    
    ' To speed things up, if the 'text' string is empty, exit the function
    If Len(text) = 0 Then Exit Function
    
    If Not (mPoFile Is Nothing) Then
        s_ = UnescapeStr(mPoFile.GetText(text))
    Else
        s_ = UnescapeStr(text)
    End If
    
    ' Replace the placeholders with parameters
    ' This code appears again in the tc() function, but I found no way
    ' to put it into a dedicated function, as it appears not to be
    ' possible to call a function with a preexisting paramarray.
    lb = LBound(params)
    ub = UBound(params)
    j = 0
    
    ' Replace {} only if there are params at all in the function call
    If lb <= ub Then
        Do
            i = InStr(s_, "{}")
            If i And (lb + j <= ub) Then
                s_ = Left$(s_, i - 1) & params(j) & Mid$(s_, i + 2)
                j = j + 1
            End If
        Loop Until i = 0
    End If ' lb <= ub
    
    ' Return the translated string
    t = s_
End Function

Function tc(ByRef context As String, ByVal text As String, ParamArray params() As Variant) As String
' Translates the text in the specified context.
' If 'context' is an empty string, this acts as the normal t() function.
    
    Dim s_ As String
    Dim i As Long
    Dim j As Long
    Dim lb As Long
    Dim ub As Long
    
    ' To speed things up, if the 'text' string is empty, exit the function
    If Len(text) = 0 Then Exit Function
    
    If Not (mPoFile Is Nothing) Then
        s_ = UnescapeStr(mPoFile.GetTextContext(context, text))
    Else
        s_ = UnescapeStr(text)
    End If
    
    ' Replace the placeholders with parameters
    ' This code appears again in the t() function, but I found no way
    ' to put it into a dedicated function, as it appears not to be
    ' possible to call a function with a preexisting paramarray.
    lb = LBound(params)
    ub = UBound(params)
    j = 0
    
    ' Replace {} only if there are params at all in the function call
    If lb <= ub Then
        Do
            i = InStr(s_, "{}")
            If i And (lb + j <= ub) Then
                s_ = Left$(s_, i - 1) & params(j) & Mid$(s_, i + 2)
                j = j + 1
            End If
        Loop Until i = 0
    End If ' lb <= ub
    
    tc = s_
End Function

Function DetectLanguage() As tLocaleInfo
' Detects the language of the current Excel installation.

    Dim i As Long
    Dim ui As Long
    
    ' If the locale info array has not been built yet, do it now.
    If mNumLocales = 0 Then CreateLocaleInfo
    
    ui = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    ' This search algorithm is very crude.
    ' One could sort the list first. But then again, there is a
    ' limited number of languages, and the user won't notice any
    ' delay.
    For i = LBound(mLocales) To UBound(mLocales)
        If mLocales(i).id = ui Then
            DetectLanguage = mLocales(i)
            Exit For
        End If
    Next i
    
End Function


Function SetLanguage(languageCode As String) As Boolean
' Attempts to load a new PO file.
' Returns True on success.
' When a PO file could not be found or loaded, the old
' mPOFile is preserved (if one was present).
    Dim po As clsGetTextFile
    Dim strCode As String
    
    ' 2013-06-02: This function now accepts language strings as returned
    ' by BuildLanguageList() to make usage more convenient.
    strCode = languageCode
    If strCode Like "*(*)*" Then
        strCode = Left$(strCode, InStr(strCode, ")") - 1)
        strCode = Mid$(strCode, InStr(strCode, "(") + 1)
    End If
    
    On Error Resume Next
    If strCode = DEFAULT_LANGUAGE Then
        mCurrentLang = strCode
        Set mPoFile = Nothing
    ElseIf strCode <> mCurrentLang Then
        If Not FileExists(GetMoFilePath(strCode)) Then
            ' If the language code is of form "en-US" and a corresponding file
            ' was not found, attempt to locate the one with the pure language
            ' code ("en" in this case") only.
            If strCode Like "??-??" Then
                strCode = Left$(strCode, 2)
                If Not FileExists(GetMoFilePath(strCode)) Then Exit Function
            Else
                Exit Function
            End If
        End If
        Set po = New clsGetTextFile
        If po.ReadFile(GetMoFilePath(strCode)) Then
            Set mPoFile = po
            mCurrentLang = strCode
            SetLanguage = True
        End If
    End If
End Function

Function GetLanguage() As String
' Returns the current language
    GetLanguage = mCurrentLang
End Function

Function PoFileExists(lang As String) As Boolean
    On Error Resume Next
    If FileExists(GetMoFilePath(lang)) Then PoFileExists = True
End Function


Function TranslateForm(f As UserForm) As Boolean
' *** COMPILE ERROR? *** If you experience compile errors with this
' function, it is because your VBA project does not contain a UserForm.
' Amend this by simply adding a UserForm to your project (even if you don't
' need a form). Or comment out this function.
'
' Goes through all elements of form f and translates their captions.
' Call this function from a Userform_Initialize constructor to
' enable L10N.
' Strangely, the userform's caption is always empty when we receive
' the form as a parameter in this function. Therefore, userforms
' need to translate their own captions in the constructor, e.g.
'     me.Caption = t(me.Caption)

    Dim c As control
    
    ' Turn off runtime exceptions, as not every control on the form
    ' will have a caption, for example.
    
    On Error Resume Next
    For Each c In f.Controls
        TranslateFormControl c
    Next c
End Function

Function EscapeStr(ByVal text As String) As String
' Escapes a string by replacing vbNewLines and vbLfs with "\n"
' and vbTabs with "\t".
    Dim i As Long
    
    ' Start with vbNewLine, which is vbCarriageReturn & vbLineFeed
    i = InStr(text, vbNewLine)
    While i
        text = Left$(text, i - 1) & "\n" & Mid$(text, i + 2)
        i = InStr(i + 2, text, vbNewLine)
    Wend
    
    i = InStr(text, vbLf)
    While i
        text = Left$(text, i - 1) & "\n" & Mid$(text, i + 1)
        i = InStr(i + 2, text, vbLf)
    Wend
    
    i = InStr(text, vbTab)
    While i
        text = Left$(text, i - 1) & "\t" & Mid$(text, i + 1)
        i = InStr(i + 2, text, vbTab)
    Wend
    
    EscapeStr = text
End Function

Function UnescapeStr(ByVal text As String) As String
' Unescapes a string by replacing \n and \t with the respective character
' bytes.
    Dim i As Long
    Dim o As Long
    Dim x As String * 2
    Dim s As String
    
    o = 0 ' The offset in the string (needed to be able to handle '\\')
    i = InStr(text, "\")
    While i
        x = Mid$(text, i, 2)
        If x = "\n" Then
            s = vbNewLine
        ElseIf x = "\t" Then
            s = vbTab
        ElseIf x = "\\" Then
            s = "\"
        Else
            s = ""
        End If
        text = Left$(text, i - 1) & s & Mid$(text, i + 2)
        o = i + 1
        i = InStr(o, text, "\")
    Wend
    UnescapeStr = text
End Function

Function EncodeAccelerator(msgId As String, accelerator As String) As String
' Encodes an accelerator key for a form control in msgId.
' The accelerator key is encoded by prepending an ampersand, e.g.
' "Check for updates" with accelerator "u" --> "Check for &updates".
' To decode the accelerator key and remove the ampersand, use the DecodeAccelerator function.
' Note that the Accelerator encoding is case-sensitive.
' If the accelerator key is not contained in the string, it will get lost, so it is up to the
' person designing a form to make sure the accelerator key is valid.
    Dim s As String
    Dim i As Integer
    s = msgId
    
    ' Escape existing ampersands
    i = InStr(s, "&")
    While i
        s = Left$(s, i - 1) & "&" & Mid$(s, i)
        i = InStr(i + 2, s, "&")
    Wend
    
    If Len(accelerator) Then
        i = InStr(s, accelerator)
        If i Then
            ' Encode the accelerator
            s = Left$(s, i - 1) & "&" & Mid$(s, i)
        End If
    End If
    
    EncodeAccelerator = s
End Function

Function DecodeAccelerator(msgStr As String, ByRef accelerator As String) As String
' Extracts an accelerator key (marked with a prefixed ampersand) from msgid.
' The accelerator key is returned in Accelerator.
' Example:
' "Check for &updates" --> "Check for updates", Accelerator:="u"
' To store an accelerator key in a MsgId, use EncodeAccelerator
    Dim i As Integer
    
    accelerator = ""
    If Len(msgStr) Then
        ' Look for a single ampersand; since pre-existing ampersands should have been
        ' escaped with another amperand by the EncodeAccelerator function when
        ' the MsgId was stored, any single ampersand must indicate the accelerator.
        i = InStr(msgStr, "&")
        Do While i
            If Mid$(msgStr, i, 2) = "&&" Then
                msgStr = Left$(msgStr, i) & Mid$(msgStr, i + 2)
                i = InStr(i + 1, msgStr, "&")
            Else
                Exit Do
            End If
        Loop
        If i Then
            accelerator = Mid$(msgStr, i + 1, 1)
            msgStr = Left$(msgStr, i - 1) & Mid$(msgStr, i + 1)
        End If
    End If
    DecodeAccelerator = msgStr
End Function

Function BuildLanguageList(ByRef currentIndex As Long) As String()
' Builds a list of available languages and returns an array of strings.
' The parameter currentIndex will contain the index of the currently
' used language.
' Since many languages are stored in the mLocales array only with the
' complete locale string (e.g., "en-US"), this function also generates
' entries for the non-localized languages (e.g., "en").
    Dim a() As String
    Dim curLang As String * 2
    Dim i As Long
    Dim n As Long
    
    ' If the list of all possible locales has not been generated
    ' yet, do it now.
    On Error Resume Next
        n = UBound(mLocales)
        If Err Then CreateLocaleInfo
    On Error GoTo 0
    
    ReDim a(LBound(mLocales) To UBound(mLocales))
    n = 0
    currentIndex = LBound(a)
    
    For i = LBound(mLocales) To UBound(mLocales)
        With mLocales(i)
            If mLocales(i).LocaleCode = DEFAULT_LANGUAGE Then
                a(LBound(a)) = .Name & " (" & .LocaleCode & ")"
                Exit For
            End If
        End With ' mLocales(i)
    Next i
    For i = LBound(mLocales) To UBound(mLocales)
        With mLocales(i)
            If .LocaleCode <> DEFAULT_LANGUAGE Then
                ' Is it a new block of languages?
                If .Code <> curLang Then
                    curLang = .Code
                    If FileExists(GetMoFilePath(.Code)) Then
                        n = n + 1
                        If UBound(a) < LBound(a) + n Then ReDim Preserve a(LBound(a) To UBound(a) + 10)
                        a(LBound(a) + n) = Left$(.Name, InStr(.Name, " -")) & " (" & .Code & ")"
                        If mCurrentLang = .Code Then currentIndex = n
                    End If
                ElseIf FileExists(GetMoFilePath(.LocaleCode)) Then
                    n = n + 1
                    If UBound(a) < LBound(a) + n Then ReDim Preserve a(LBound(a) To UBound(a) + 10)
                    a(LBound(a) + n) = .Name & " (" & .Code & ")"
                    If mCurrentLang = .LocaleCode Then currentIndex = n
                End If
            End If
        End With ' mLocales(i)
    Next i
    ReDim Preserve a(LBound(a) To LBound(a) + n)
    BuildLanguageList = a
End Function


' =================================================================
' == Private module functions
' =================================================================


Private Sub TranslateFormControl(c As control)
' Translates the caption and tool tip text of the control.
' Calls itself recursively if the control is a container (such as
' a Frame or MultiPage).
    Dim children As Controls
    Dim child As control
    Dim accel As String
    Dim p As MSForms.Page
    On Error Resume Next
    With c
        .ControlTipText = t(.ControlTipText)
        
        ' The following complex line needs explanation:
        ' First, the EncodeAccelerator is called with the Caption and Accelerator
        ' to get the format in which the MsgId is stored in the PO/MO file.
        ' Then, we look up the MsgId using the t() function. This will return a
        ' MsgStr which may contain an encoded accelerator and escaped ampersands ("&&"),
        ' which are subsequently decoded by the DecodeAccelerator function.
        ' The EscapeStr() function is called to convert newline characters and the like.
        accel = .accelerator
        .Caption = DecodeAccelerator(t(EscapeStr(EncodeAccelerator(.Caption, accel))), accel)
        .accelerator = accel
        
        Set children = c.Controls
        If Not (children Is Nothing) Then
            For Each child In children
                TranslateFormControl child
            Next child
        End If
        
        ' Multipage controls are a special case
        If TypeName(c) = "MultiPage" Then
            For Each p In c.Pages
                p.Caption = t(p.Caption)
            Next p
        End If
    End With 'c
End Sub

Private Function GetMoFilePath(lang As String) As String
' Constructs the path to the language file
    Dim path As String
    
    path = AddPathSep(ThisWorkbook.path)
    GetMoFilePath = AddPathSep(path & LOCALE_DIR) & _
        LOCALE_BASE_NAME & "_" & lang & ".mo"
    
End Function


Private Sub AddLangInfo(LanguageName, Code, LocaleCode, id)
' Helper procedure that populates the module-private array with locale
' info.

    With mLocales(mNumLocales)
        .Name = LanguageName
        
        ' Make sure that the locale code capitalization follows
        ' the pattern "xx-XX".
        If LocaleCode Like "??-??" Then
            LocaleCode = Left$(LocaleCode, 3) & UCase$(Mid$(LocaleCode, 4))
        End If
        .LocaleCode = LocaleCode
        .Code = Code
        .id = id
    End With ' mLocales
    mNumLocales = mNumLocales + 1
End Sub

Private Sub CreateLocaleInfo()
' Creates an array with locale info. Since Excel VBA cannot deal
' with array constants, we use a work around, evoking a function.
' The commented block of code is a list of pretty much all the languages.
' We only use the languages that we actually have translations for,
' and use the localized forms of the language names.
    
    ' Clear the mLocales array.
    ReDim mLocales(0 To 189) As tLocaleInfo
    
    AddLangInfo "Deutsch - Deutschland", "de", "de-de", 1031
    AddLangInfo "English - United States", "en", "en-us", 1033
   ' AddLangInfo "Nederlands - Nederland", "nl", "nl-nl", 1043
    
    
'    AddLangInfo "Afrikaans", "af", "af", 1078
'    AddLangInfo "Albanian", "sq", "sq", 1052
'    AddLangInfo "Amharic", "am", "am", 1118
'    AddLangInfo "Arabic - Algeria", "ar", "ar-dz", 5121
'    AddLangInfo "Arabic - Bahrain", "ar", "ar-bh", 15361
'    AddLangInfo "Arabic - Egypt", "ar", "ar-eg", 3073
'    AddLangInfo "Arabic - Iraq", "ar", "ar-iq", 2049
'    AddLangInfo "Arabic - Jordan", "ar", "ar-jo", 11265
'    AddLangInfo "Arabic - Kuwait", "ar", "ar-kw", 13313
'    AddLangInfo "Arabic - Lebanon", "ar", "ar-lb", 12289
'    AddLangInfo "Arabic - Libya", "ar", "ar-ly", 4097
'    AddLangInfo "Arabic - Morocco", "ar", "ar-ma", 6145
'    AddLangInfo "Arabic - Oman", "ar", "ar-om", 8193
'    AddLangInfo "Arabic - Qatar", "ar", "ar-qa", 16385
'    AddLangInfo "Arabic - Saudi Arabia", "ar", "ar-sa", 1025
'    AddLangInfo "Arabic - Syria", "ar", "ar-sy", 10241
'    AddLangInfo "Arabic - Tunisia", "ar", "ar-tn", 7169
'    AddLangInfo "Arabic - United Arab Emirates", "ar", "ar-ae", 14337
'    AddLangInfo "Arabic - Yemen", "ar", "ar-ye", 9217
'    AddLangInfo "Armenian", "hy", "hy", 1067
'    AddLangInfo "Assamese", "as", "as", 1101
'    AddLangInfo "Azeri - Cyrillic", "az", "az-az", 2092
'    AddLangInfo "Azeri - Latin", "az", "az-az", 1068
'    AddLangInfo "Basque", "eu", "eu", 1069
'    AddLangInfo "Belarusian", "be", "be", 1059
'    AddLangInfo "Bengali - Bangladesh", "bn", "bn", 2117
'    AddLangInfo "Bengali - India", "bn", "bn", 1093
'    AddLangInfo "Bosnian", "bs", "bs", 5146
'    AddLangInfo "Bulgarian", "bg", "bg", 1026
'    AddLangInfo "Burmese", "my", "my", 1109
'    AddLangInfo "Catalan", "ca", "ca", 1027
'    AddLangInfo "Chinese - China", "zh", "zh-cn", 2052
'    AddLangInfo "Chinese - Hong Kong SAR", "zh", "zh-hk", 3076
'    AddLangInfo "Chinese - Macau SAR", "zh", "zh-mo", 5124
'    AddLangInfo "Chinese - Singapore", "zh", "zh-sg", 4100
'    AddLangInfo "Chinese - Taiwan", "zh", "zh-tw", 1028
'    AddLangInfo "Croatian", "hr", "hr", 1050
'    AddLangInfo "Czech", "cs", "cs", 1029
'    AddLangInfo "Danish", "da", "da", 1030
'    AddLangInfo "Dutch - Belgium", "nl", "nl-be", 2067
'    AddLangInfo "Dutch - Netherlands", "nl", "nl-nl", 1043
'    AddLangInfo "Edo", "", "", 1126
'    AddLangInfo "English - Australia", "en", "en-au", 3081
'    AddLangInfo "English - Belize", "en", "en-bz", 10249
'    AddLangInfo "English - Canada", "en", "en-ca", 4105
'    AddLangInfo "English - Caribbean", "en", "en-cb", 9225
'    AddLangInfo "English - Great Britain", "en", "en-gb", 2057
'    AddLangInfo "English - India", "en", "en-in", 16393
'    AddLangInfo "English - Ireland", "en", "en-ie", 6153
'    AddLangInfo "English - Jamaica", "en", "en-jm", 8201
'    AddLangInfo "English - New Zealand", "en", "en-nz", 5129
'    AddLangInfo "English - Phillippines", "en", "en-ph", 13321
'    AddLangInfo "English - Southern Africa", "en", "en-za", 7177
'    AddLangInfo "English - Trinidad", "en", "en-tt", 11273
'    AddLangInfo "English - United States", "en", "en-us", 1033
'    AddLangInfo "English - Zimbabwe", "en", "", 12297
'    AddLangInfo "Estonian", "et", "et", 1061
'    AddLangInfo "Faroese", "fo", "fo", 1080
'    AddLangInfo "Farsi - Persian", "fa", "fa", 1065
'    AddLangInfo "Filipino", "", "", 1124
'    AddLangInfo "Finnish", "fi", "fi", 1035
'    AddLangInfo "French - Belgium", "fr", "fr-be", 2060
'    AddLangInfo "French - Cameroon", "fr", "", 11276
'    AddLangInfo "French - Canada", "fr", "fr-ca", 3084
'    AddLangInfo "French - Congo", "fr", "", 9228
'    AddLangInfo "French - Cote d'Ivoire", "fr", "", 12300
'    AddLangInfo "French - France", "fr", "fr-fr", 1036
'    AddLangInfo "French - Luxembourg", "fr", "fr-lu", 5132
'    AddLangInfo "French - Mali", "fr", "", 13324
'    AddLangInfo "French - Monaco", "fr", "", 6156
'    AddLangInfo "French - Morocco", "fr", "", 14348
'    AddLangInfo "French - Senegal", "fr", "", 10252
'    AddLangInfo "French - Switzerland", "fr", "fr-ch", 4108
'    AddLangInfo "French - West Indies", "fr", "", 7180
'    AddLangInfo "Frisian - Netherlands", "", "", 1122
'    AddLangInfo "FYRO Macedonia", "mk", "mk", 1071
'    AddLangInfo "Gaelic - Ireland", "gd", "gd-ie", 2108
'    AddLangInfo "Gaelic - Scotland", "gd", "gd", 1084
'    AddLangInfo "Galician", "gl", "", 1110
'    AddLangInfo "Georgian", "ka", "", 1079
'    AddLangInfo "German - Austria", "de", "de-at", 3079
'    AddLangInfo "German - Germany", "de", "de-de", 1031
'    AddLangInfo "German - Liechtenstein", "de", "de-li", 5127
'    AddLangInfo "German - Luxembourg", "de", "de-lu", 4103
'    AddLangInfo "German - Switzerland", "de", "de-ch", 2055
'    AddLangInfo "Greek", "el", "el", 1032
'    AddLangInfo "Guarani - Paraguay", "gn", "gn", 1140
'    AddLangInfo "Gujarati", "gu", "gu", 1095
'    AddLangInfo "Hebrew", "he", "he", 1037
'    AddLangInfo "HID  Human Interface Device", "", "", 1279
'    AddLangInfo "Hindi", "hi", "hi", 1081
'    AddLangInfo "Hungarian", "hu", "hu", 1038
'    AddLangInfo "Icelandic", "is", "is", 1039
'    AddLangInfo "Igbo - Nigeria", "", "", 1136
'    AddLangInfo "Indonesian", "id", "id", 1057
'    AddLangInfo "Italian - Italy", "it", "it-it", 1040
'    AddLangInfo "Italian - Switzerland", "it", "it-ch", 2064
'    AddLangInfo "Japanese", "ja", "ja", 1041
'    AddLangInfo "Kannada", "kn", "kn", 1099
'    AddLangInfo "Kashmiri", "ks", "ks", 1120
'    AddLangInfo "Kazakh", "kk", "kk", 1087
'    AddLangInfo "Khmer", "km", "km", 1107
'    AddLangInfo "Konkani", "", "", 1111
'    AddLangInfo "Korean", "ko", "ko", 1042
'    AddLangInfo "Kyrgyz - Cyrillic", "", "", 1088
'    AddLangInfo "Lao", "lo", "lo", 1108
'    AddLangInfo "Latin", "la", "la", 1142
'    AddLangInfo "Latvian", "lv", "lv", 1062
'    AddLangInfo "Lithuanian", "lt", "lt", 1063
'    AddLangInfo "Malay - Brunei", "ms", "ms-bn", 2110
'    AddLangInfo "Malay - Malaysia", "ms", "ms-my", 1086
'    AddLangInfo "Malayalam", "ml", "ml", 1100
'    AddLangInfo "Maltese", "mt", "mt", 1082
'    AddLangInfo "Manipuri", "", "", 1112
'    AddLangInfo "Maori", "mi", "mi", 1153
'    AddLangInfo "Marathi", "mr", "mr", 1102
'    AddLangInfo "Mongolian", "mn", "mn", 2128
'    AddLangInfo "Mongolian", "mn", "mn", 1104
'    AddLangInfo "Nepali", "ne", "ne", 1121
'    AddLangInfo "Norwegian - Bokml", "nb", "no-no", 1044
'    AddLangInfo "Norwegian - Nynorsk", "nn", "no-no", 2068
'    AddLangInfo "Oriya", "or", "or", 1096
'    AddLangInfo "Polish", "pl", "pl", 1045
'    AddLangInfo "Portuguese - Brazil", "pt", "pt-br", 1046
'    AddLangInfo "Portuguese - Portugal", "pt", "pt-pt", 2070
'    AddLangInfo "Punjabi", "pa", "pa", 1094
'    AddLangInfo "Raeto-Romance", "rm", "rm", 1047
'    AddLangInfo "Romanian - Moldova", "ro", "ro-mo", 2072
'    AddLangInfo "Romanian - Romania", "ro", "ro", 1048
'    AddLangInfo "Russian", "ru", "ru", 1049
'    AddLangInfo "Russian - Moldova", "ru", "ru-mo", 2073
'    AddLangInfo "Sami Lappish", "", "", 1083
'    AddLangInfo "Sanskrit", "sa", "sa", 1103
'    AddLangInfo "Serbian - Cyrillic", "sr", "sr-sp", 3098
'    AddLangInfo "Serbian - Latin", "sr", "sr-sp", 2074
'    AddLangInfo "Sesotho  Sutu", "", "", 1072
'    AddLangInfo "Setsuana", "tn", "tn", 1074
'    AddLangInfo "Sindhi", "sd", "sd", 1113
'    AddLangInfo "Sinhala", "si", "si", 1115
'    AddLangInfo "Slovak", "sk", "sk", 1051
'    AddLangInfo "Slovenian", "sl", "sl", 1060
'    AddLangInfo "Somali", "so", "so", 1143
'    AddLangInfo "Sorbian", "sb", "sb", 1070
'    AddLangInfo "Spanish - Argentina", "es", "es-ar", 11274
'    AddLangInfo "Spanish - Bolivia", "es", "es-bo", 16394
'    AddLangInfo "Spanish - Chile", "es", "es-cl", 13322
'    AddLangInfo "Spanish - Colombia", "es", "es-co", 9226
'    AddLangInfo "Spanish - Costa Rica", "es", "es-cr", 5130
'    AddLangInfo "Spanish - Dominican Republic", "es", "es-do", 7178
'    AddLangInfo "Spanish - Ecuador", "es", "es-ec", 12298
'    AddLangInfo "Spanish - El Salvador", "es", "es-sv", 17418
'    AddLangInfo "Spanish - Guatemala", "es", "es-gt", 4106
'    AddLangInfo "Spanish - Honduras", "es", "es-hn", 18442
'    AddLangInfo "Spanish - Mexico", "es", "es-mx", 2058
'    AddLangInfo "Spanish - Nicaragua", "es", "es-ni", 19466
'    AddLangInfo "Spanish - Panama", "es", "es-pa", 6154
'    AddLangInfo "Spanish - Paraguay", "es", "es-py", 15370
'    AddLangInfo "Spanish - Peru", "es", "es-pe", 10250
'    AddLangInfo "Spanish - Puerto Rico", "es", "es-pr", 20490
'    AddLangInfo "Spanish - Spain  Traditional", "es", "es-es", 1034
'    AddLangInfo "Spanish - Uruguay", "es", "es-uy", 14346
'    AddLangInfo "Spanish - Venezuela", "es", "es-ve", 8202
'    AddLangInfo "Swahili", "sw", "sw", 1089
'    AddLangInfo "Swedish - Finland", "sv", "sv-fi", 2077
'    AddLangInfo "Swedish - Sweden", "sv", "sv-se", 1053
'    AddLangInfo "Syriac", "", "", 1114
'    AddLangInfo "Tajik", "tg", "tg", 1064
'    AddLangInfo "Tamil", "ta", "ta", 1097
'    AddLangInfo "Tatar", "tt", "tt", 1092
'    AddLangInfo "Telugu", "te", "te", 1098
'    AddLangInfo "Thai", "th", "th", 1054
'    AddLangInfo "Tibetan", "bo", "bo", 1105
'    AddLangInfo "Tsonga", "ts", "ts", 1073
'    AddLangInfo "Turkish", "tr", "tr", 1055
'    AddLangInfo "Turkmen", "tk", "tk", 1090
'    AddLangInfo "Ukrainian", "uk", "uk", 1058
'    AddLangInfo "Unicode", "", "UTF-8", 0
'    AddLangInfo "Urdu", "ur", "ur", 1056
'    AddLangInfo "Uzbek - Cyrillic", "uz", "uz-uz", 2115
'    AddLangInfo "Uzbek - Latin", "uz", "uz-uz", 1091
'    AddLangInfo "Venda", "", "", 1075
'    AddLangInfo "Vietnamese", "vi", "vi", 1066
'    AddLangInfo "Welsh", "cy", "cy", 1106
'    AddLangInfo "Xhosa", "xh", "xh", 1076
'    AddLangInfo "Yiddish", "yi", "yi", 1085
'    AddLangInfo "Zulu", "zu", "zu", 1077

End Sub

Function ComputeHash(str As String) As Long
    ComputeHash = compute_hashval(str, Len(str))
End Function
