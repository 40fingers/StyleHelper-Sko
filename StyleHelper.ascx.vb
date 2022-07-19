Option Strict On

Imports System.Text
Imports System.Globalization
Imports System.Web.UI.WebControls
Imports System.Collections.Generic

Imports DotNetNuke
Imports DotNetNuke.Security.Permissions
Imports Dotnetnuke.Entities.Modules


Imports DotNetNuke.Application


Namespace FortyFingers.Dnn.SkinObjects

    ''' ----------------------------------------------------------------------------- 
    ''' <summary></summary> 
    ''' <remarks></remarks> 
    ''' <history> 
    ''' <todo>
    ''' </todo>
    ''' 
    ''' </history> 
    ''' ----------------------------------------------------------------------------- 
    ''' 
    Partial Class StyleHelper
        Inherits UI.Skins.SkinObjectBase


#Region "Private Members"

        Private Const _sNOT As String = "!"
        Private Const _strModuleEdit As String = "EditModule"

        'Check if the conditions are met to proceed
        Dim bConditions As Boolean

        Enum InjectPosition
            Top = 0
            Bottom = 1
        End Enum

        Enum TextType

            Text = 0
            CssClass = 1
            Id = 2
            EncodedHtml = 3

        End Enum


        Enum CssClassCases

            None = 0
            lowercase = 1
            UPPERCASE = 2
            PascalCase = 3

        End Enum

        Enum VersionIs

            Same = 0
            Bigger = 1
            Smaller = -1

        End Enum


        Dim strUserModeNone As String = "None"

        Dim caCondition() As Char = "<>=".ToCharArray()



#End Region


#Region "Public Properties Add & Remove"


        Private _AddToHead As String
        ''' <summary>
        ''' Pipe Sperated strings to add to the head of the page
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddToHead() As String
            Get
                Return _AddToHead
            End Get
            Set(ByVal value As String)
                _AddToHead = value
            End Set
        End Property


        Private _RemoveFromHead As String
        ''' <summary>
        ''' Pipe Sperated strings to remove to the head of the page
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RemoveFromHead() As String
            Get
                Return _RemoveFromHead
            End Get
            Set(ByVal value As String)
                _RemoveFromHead = value
            End Set
        End Property



        Private _RemoveMeta As String
        ''' <summary>
        ''' Pipe seperated ids of controle to disable
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RemoveMeta() As String
            Get
                Return _RemoveMeta
            End Get
            Set(ByVal value As String)
                _RemoveMeta = value
            End Set
        End Property



        Private _ChangeMeta As String
        ''' <summary>
        ''' Pipe Separated ids to Parse
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ChangeMeta() As String
            Get
                Return _ChangeMeta
            End Get
            Set(ByVal value As String)
                _ChangeMeta = value
            End Set
        End Property


        Private _MetaNameToLower As Boolean
        ''' <summary>
        ''' Set the DNN core meta tag name attributes to lowercase
        ''' </summary>
        ''' <returns></returns>
        Public Property MetaNameToLower() As Boolean
            Get
                Return _MetaNameToLower
            End Get
            Set(ByVal value As Boolean)
                _MetaNameToLower = value
            End Set
        End Property



        Private _sRemoveCssFile As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of CSS files to remove from the head of the page
        ''' </summary>
        ''' <value>Empty</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RemoveCssFile() As String
            Set(ByVal Value As String)
                _sRemoveCssFile = Value
            End Set
            Get
                Return _sRemoveCssFile
            End Get

        End Property


        Private _sRemoveControls As String = String.Empty
        ''' <summary>
        ''' Remove any control (comma seperated)
        ''' </summary>
        ''' <remarks></remarks>
        Public Property RemoveControls() As String
            Get
                Return _sRemoveControls
            End Get
            Set(ByVal value As String)
                _sRemoveControls = value
            End Set
        End Property





        Private _bFilterRemove As Boolean = True
        ''' <summary>
        ''' Set if the removal of CSS files should be conditional
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilterRemove() As Boolean
            Set(ByVal Value As Boolean)
                _bFilterRemove = Value
            End Set
            Get
                Return _bFilterRemove
            End Get
        End Property



        Private _sAddCssFile As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of CSS files to add
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddCssFile() As String
            Set(ByVal Value As String)
                _sAddCssFile = Value
            End Set
            Get
                Return _sAddCssFile
            End Get
        End Property



        Private _sAddJsFile As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of JS files to add
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddJsFile() As String
            Set(ByVal Value As String)
                _sAddJsFile = Value
            End Set
            Get
                Return _sAddJsFile
            End Get

        End Property



        Private _sRemoveJsFile As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of JS files to remove from the Page (only the ones added via DNN are supported)
        ''' </summary>
        ''' <value>Empty</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RemoveJsFile() As String
            Set(ByVal Value As String)
                _sRemoveJsFile = Value
            End Set
            Get
                Return _sRemoveJsFile
            End Get

        End Property



        Private _bFilterAdd As Boolean = True
        ''' <summary>
        ''' Set if the adding should be conditional
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilterAdd() As Boolean
            Set(ByVal Value As Boolean)
                _bFilterAdd = Value
            End Set
            Get
                Return _bFilterAdd
            End Get
        End Property



        Private _sAddMetaTags As String = String.Empty
        ''' <summary>
        ''' List of Meta tags to add to the head.
        ''' Format: [Name:Content]
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddMetaTags() As String
            Set(ByVal Value As String)
                _sAddMetaTags = Value
            End Set
            Get
                Return _sAddMetaTags
            End Get

        End Property



        Private _bFilterMeta As Boolean = True
        ''' <summary>
        ''' Set if the removal of CSS files should be conditional
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilterMeta() As Boolean
            Set(ByVal Value As Boolean)
                _bFilterMeta = Value
            End Set
            Get
                Return _bFilterMeta
            End Get
        End Property



        Private _sBasePath As String = "[S]"
        ''' <summary>
        ''' Add skinpath to paths
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BasePath() As String
            Set(ByVal Value As String)
                _sBasePath = Value
            End Set
            Get
                Return _sBasePath
            End Get
        End Property



        Private _bAddAtEnd As Boolean = True
        ''' <summary>
        ''' Add links at the end of the page head
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddAtEnd() As Boolean
            Set(ByVal Value As Boolean)
                _bAddAtEnd = Value
            End Set
            Get
                Return _bAddAtEnd
            End Get
        End Property



        Private _bCorrectModuleCss As Boolean
        ''' <summary>
        ''' Correct Module.css load order
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CorrectModuleCss() As Boolean
            Get
                Return _bCorrectModuleCss
            End Get
            Set(ByVal value As Boolean)
                _bCorrectModuleCss = value
            End Set
        End Property


        Private _bForceCssOrder As Boolean
        ''' <summary>
        ''' Forces Skin.css to be loaded beforelast and portal.css last.
        ''' </summary>
        ''' <returns></returns>
        Public Property ForceSkinCssOrder() As Boolean
            Get
                Return _bForceCssOrder
            End Get
            Set(ByVal value As Boolean)
                _bForceCssOrder = value
            End Set
        End Property




        Private _sCssMedia As String = "screen"
        ''' <summary>
        ''' Media type for CSS files
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property CssMedia() As String
            Set(ByVal Value As String)
                _sCssMedia = Value
            End Set
            Get
                Return _sCssMedia
            End Get
        End Property





        Private _bShowInfo As Boolean
        ''' <summary>
        ''' Render ShowInfo info in browser, like browser version
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property ShowInfo() As Boolean
            Get
                Return _bShowInfo
            End Get
            Set(ByVal Value As Boolean)
                _bShowInfo = Value
            End Set
        End Property




        Dim _bAddBodyClass As Boolean = False
        ''' <summary>
        ''' Add Page Class to Body Element without filtering
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddBodyClass() As Boolean
            Set(ByVal Value As Boolean)
                _bAddBodyClass = Value
            End Set
            Get
                Return _bAddBodyClass
            End Get

        End Property



        ''' <summary>
        ''' Legacy: replace with AddToBodyTop
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddAfterBody() As String
            Get
                Return AddToBodyTop
            End Get
            Set(ByVal value As String)
                AddToBodyTop = value
            End Set
        End Property



        Private _sAddToBodyTop As String
        ''' <summary>
        ''' Add at the top of the body element
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddToBodyTop() As String
            Get
                Return _sAddToBodyTop
            End Get
            Set(ByVal value As String)
                _sAddToBodyTop = value
            End Set
        End Property


        Private _sAddToBodyBottom As String
        ''' <summary>
        ''' Add at the bottom of the body element
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AddToBodyBottom() As String
            Get
                Return _sAddToBodyBottom
            End Get
            Set(ByVal value As String)
                _sAddToBodyBottom = value
            End Set
        End Property






        Dim _bFilterBodyClass As Boolean = False
        ''' <summary>
        ''' Filter adding the body class or not (by default false)?
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilterBodyClass() As Boolean
            Set(ByVal Value As Boolean)
                _bFilterBodyClass = Value
            End Set
            Get
                Return _bFilterBodyClass
            End Get

        End Property




        Dim _sBodyClass As String = "Page-[Page:Name] Level-[Page:Level] [BcName] [BcId] [BcNr] CP-[CPState] [PageType] [UserPageRoles] Cult-[Culture] Lang-[Language] [IE]"

        ''' <summary>
        ''' Template for the page class
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BodyClass() As String
            Get
                Return (_sBodyClass)
            End Get
            Set(ByVal value As String)
                'If the template is set, set adding the class to true.
                AddBodyClass = True
                _sBodyClass = value
            End Set
        End Property


        Dim _sAddToBodyClass As String = String.Empty
        ''' <summary>
        ''' Add something to the existing / prevously set body class, used when you insert a class conditinally.
        ''' </summary>
        ''' <returns></returns>

        Public Property AddToBodyClass() As String
            Get
                Return (_sAddToBodyClass)
            End Get
            Set(ByVal value As String)
                'If the template is set, set adding the class to true.
                AddBodyClass = True
                _sAddToBodyClass = value
            End Set
        End Property



        Private _sCssClassCase As String = "none"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        Public Property CssClassCase() As String
            Get
                Return _sCssClassCase
            End Get
            Set(ByVal value As String)
                _sCssClassCase = value
            End Set
        End Property



        Private _sAddHTMLAttribte As String
        ''' <summary>
        ''' Write an attribute to the HTML element
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        Public Property AddHtmlAttribute() As String
            Get
                Return _sAddHTMLAttribte
            End Get
            Set(ByVal value As String)
                _sAddHTMLAttribte = value
            End Set
        End Property




        Private _sDocType As String
        ''' <summary>
        ''' Set the skin doctype
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        Public Property Doctype() As String
            Get
                Return _sDocType
            End Get
            Set(ByVal value As String)
                _sDocType = value
            End Set
        End Property



        Private _sContent As String = "<!--40Fingers Stylehelper Conditions Return True-->"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Content() As String
            Get
                Return _sContent
            End Get
            Set(ByVal value As String)
                _sContent = value
            End Set
        End Property


        Private _sContentFalse As String = "<!--40Fingers Stylehelper Conditions Return False -->"
        ''' <summary>
        ''' Content to inject when the criteria are not met
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ContentFalse() As String
            Get
                Return _sContentFalse
            End Get
            Set(ByVal value As String)
                _sContentFalse = value
            End Set
        End Property



        Private _sSplitChar As String = "||"
        ''' <summary>
        ''' Character(s) to use to pass lists
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SplitChar() As String
            Get
                Return _sSplitChar
            End Get
            Set(ByVal value As String)
                _sSplitChar = value
            End Set
        End Property






#End Region

#Region "Ml properties"

        Private _sLanguageLinksTemplate As String = "<link rel='alternate' hreflang='[LOCALE:CODE]' href='[URL]' />"


        Public Property LanguageLinksTemplate() As String
            Get
                Return _sLanguageLinksTemplate
            End Get
            Set(ByVal value As String)
                _sLanguageLinksTemplate = value
            End Set
        End Property



        Private _bAddLanguageLinksToHead As Boolean
        ''' <summary>
        ''' Adds Language links head for the other languages thi spage is available in
        ''' </summary>
        ''' <returns></returns>
        Public Property AddLanguageLinksToHead() As Boolean
            Get
                Return _bAddLanguageLinksToHead
            End Get
            Set(ByVal value As Boolean)
                _bAddLanguageLinksToHead = value
            End Set
        End Property



        Private _bFilterAddLanguageLinksToHead As Boolean = False
        ''' <summary>
        ''' Set if AddLanguageLinksToHead should check filters
        ''' </summary>
        ''' <value>False</value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FilterAddLanguageLinksToHead() As Boolean
            Set(ByVal Value As Boolean)
                _bFilterAddLanguageLinksToHead = Value
            End Set
            Get
                Return _bFilterAddLanguageLinksToHead
            End Get
        End Property

#End Region


#Region "Public Properties Filters"


        Private _strIfDnnVersion As String = String.Empty

        ''' <summary>
        ''' Check the DNN version
        ''' </summary>
        ''' <returns></returns>
        Public Property IfDnnVersion() As String
            Get
                Return _strIfDnnVersion
            End Get
            Set(ByVal value As String)
                _strIfDnnVersion = value
            End Set
        End Property


        Private _sIfBrowser As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of browsers / versions
        ''' </summary>
        ''' <value>
        ''' IE, IE10, IE>10, IE=>11
        ''' </value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfBrowser() As String
            Set(ByVal value As String)
                _sIfBrowser = value
            End Set
            Get
                Return _sIfBrowser
            End Get
        End Property




        Private _iIfMobile As Integer = 0
        ''' <summary>
        ''' If Is Mobile
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfMobile() As Boolean
            Set(ByVal value As Boolean)
                If value Then
                    _iIfMobile = 1
                Else
                    _iIfMobile = -1
                End If
            End Set
            Get
                If _iIfMobile = 1 Then
                    IfMobile = True
                Else
                    IfMobile = False
                End If
            End Get
        End Property




        Private _sIfUserAgentString As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of UserAgent Strings (regex)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfUserAgentString() As String
            Set(ByVal value As String)
                _sIfUserAgentString = value
            End Set
            Get
                Return _sIfUserAgentString
            End Get
        End Property



        ''' <summary>
        ''' Legacy Support: Replaced by IfUserAgentString
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfUserNameAgentString() As String
            Set(ByVal value As String)
                _sIfUserAgentString = value
            End Set
            Get
                Return _sIfUserAgentString
            End Get
        End Property




        Private _sIfUserName As String = String.Empty
        ''' <summary>
        ''' Comma seperated list of browsers / versions
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfUserName() As String
            Set(ByVal value As String)
                _sIfUserName = value
            End Set
            Get
                Return _sIfUserName
            End Get
        End Property




        Private _sIfUserMode As String
        ''' <summary>
        ''' If the user mode is "None,View,Edit,Layout, EditModule'
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfUserMode() As String
            Get
                Return _sIfUserMode
            End Get
            Set(ByVal value As String)
                _sIfUserMode = value
            End Set
        End Property




        Private _sIfRole As String = String.Empty
        ''' <summary>
        ''' Comma seperated list roles the user is in
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfRole() As String
            Set(ByVal value As String)
                _sIfRole = value
            End Set
            Get
                Return _sIfRole
            End Get
        End Property




        Private _bNoViewRights As Boolean = False
        ''' <summary>
        ''' When the user does not have view rights to the page. Can be used to redirect them to another page...
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfNoViewRights() As Boolean
            Set(ByVal value As Boolean)
                _bNoViewRights = value
            End Set
            Get
                Return _bNoViewRights
            End Get
        End Property



        Private _sIfURL As String = String.Empty
        ''' <summary>
        ''' Use the current pages URL for filtering
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfURL() As String
            Get
                Return _sIfURL
            End Get
            Set(ByVal value As String)
                _sIfURL = value
            End Set
        End Property



        Private _sIfCulture As String = String.Empty
        ''' <summary>
        ''' Cultures to process file for
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>

        Public Property IfCulture() As String
            Set(ByVal Value As String)
                _sIfCulture = Value
            End Set
            Get
                Return _sIfCulture
            End Get
        End Property


        Private _strIfToken As String = String.Empty
        ''' <summary>
        ''' If this is empty or all Token passed return more than an empty String this returns True 
        ''' </summary>
        ''' <returns></returns>
        Public Property IfToken() As String
            Get
                Return _strIfToken
            End Get
            Set(ByVal value As String)
                _strIfToken = value
            End Set
        End Property



        Private _bIfExternal As Boolean = True
        ''' <summary>
        ''' You can pass an external parameter from the skin by using IfExternal="<%# True %>" / IfExternal="<%# SomeFunction() %>"
        ''' </summary>
        ''' <returns></returns>
        Public Property IfExternal() As String
            Get
                Return _bIfExternal.ToString
            End Get
            Set(ByVal value As String)
                _bIfExternal = CBool(value)

            End Set
        End Property



        Private _sIfTxtDir As String = String.Empty
        ''' <summary>
        ''' Render accoring to the text direction
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfTextDir() As String
            Get
                Return (_sIfTxtDir)
            End Get
            Set(ByVal value As String)
                _sIfTxtDir = value.ToLower
            End Set
        End Property



        Private _sIfQS As String = String.Empty
        ''' <summary>
        ''' Render accoring to query string parameter
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfQS() As String
            Get
                Return (_sIfQS)
            End Get
            Set(ByVal value As String)
                _sIfQS = value.ToLower
            End Set
        End Property



        Private _sIfCookie As String
        ''' <summary>
        ''' Condition on Cookie
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfCookie() As String
            Get
                Return _sIfCookie
            End Get
            Set(ByVal value As String)
                _sIfCookie = value
            End Set
        End Property



        Private _sIfNoCookie As String
        ''' <summary>
        ''' Condition on Cookie Existance
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property IfNoCookie() As String
            Get
                Return _sIfNoCookie
            End Get
            Set(ByVal value As String)
                _sIfNoCookie = value
            End Set
        End Property



#End Region

#Region "Public Properties Info"



        Private _strTest As String = "Test"
        Public Property Test() As String
            Get
                Return _strTest
            End Get
            Set(ByVal value As String)
                _strTest = value
            End Set
        End Property
		
		
		Public ReadOnly Property PortalAlias() As String
            Get
                Return String.Format("{0}://{1}", GetAliasProtocol() , PortalSettings.PortalAlias.HTTPAlias)
            End Get
        End Property


        Public ReadOnly Property Portal() As PortalInfo
            Get
                Dim _p As New PortalInfo

                Return _p

            End Get
        End Property


#End Region


#Region "Client Detect Properties"

        Private _sClientDetectMethod As String = "detectmobilebrowsers.com"
        'For later use

        Private _sDetectMobileRegex1 As String = "(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile.+firefox|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows (ce|phone)|xda|xiino"
        ''' <summary>
        ''' First regular expression to detect Mobile browsers
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DetectMobileRegex1() As String
            Set(ByVal value As String)
                _sDetectMobileRegex1 = value
            End Set
            Get
                Return _sDetectMobileRegex1
            End Get
        End Property


        Private _bDetectMobileIncludeTablet As Boolean
        ''' <summary>
        ''' If st to true Tablets are treated as Mobile following the method on http://detectmobilebrowsers.com/about
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property DetectMobileIncludeTablet() As Boolean
            Get
                Return _bDetectMobileIncludeTablet
            End Get
            Set(ByVal value As Boolean)
                If value Then DetectMobileRegex1 = DetectMobileRegex1 & "|android|ipad|playbook|silk"
                _bDetectMobileIncludeTablet = value

            End Set
        End Property



        Private _sDetectMobileRegex2 As String = "1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-"
        ''' <summary>
        ''' First regular expression to detect Mobile browsers
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DetectMobileRegex2() As String
            Set(ByVal value As String)
                _sDetectMobileRegex2 = value
            End Set
            Get
                Return _sDetectMobileRegex2
            End Get
        End Property

#End Region


#Region "Redirect Properties"

        Private _sRedirectName As String = "Default"
        ''' <summary>
        ''' always, once, session
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property RedirectName() As String
            Get
                Return _sRedirectName
            End Get
            Set(ByVal value As String)
                _sRedirectName = value
            End Set
        End Property



        Private _sRedirectMode As String = "Session"
        ''' <summary>
        ''' always, once, session, never
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property RedirectMode() As String
            Get
                Return _sRedirectMode
            End Get
            Set(ByVal value As String)
                _sRedirectMode = value
            End Set
        End Property



        Private _sRedirectTo As String
        ''' <summary>
        ''' Redirect to url, page.name
        ''' </summary>
        ''' <remarks></remarks>
        Public Property RedirectTo() As String
            Get
                Return _sRedirectTo
            End Get
            Set(ByVal value As String)
                _sRedirectTo = value
            End Set
        End Property




        Private _sRedirectStop As String = "RedirectUrl,QueryString"
        ''' <summary>
        ''' One what parameters to refuse the redirect...
        ''' Revisit = Will refuse the redirect on second visit
        ''' RedirectUrl = Referrer is the Redirect URL
        ''' BaseRedirectUrl = Referrer contains the Base Redirect URL
        ''' QueryString = QS parameter passed 
        ''' Never = Never refuse the redirect
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property RedirectStop() As String
            Get
                Return _sRedirectStop
            End Get
            Set(ByVal value As String)
                _sRedirectStop = value
            End Set
        End Property



        Private _sRedirectBaseUrl As String = String.Empty
        ''' <summary>
        ''' The base URL for redirecting, will be used to compare
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public Property RedirectBaseUrl() As String
            Get
                Return _sRedirectBaseUrl
            End Get
            Set(ByVal value As String)
                _sRedirectBaseUrl = value
            End Set
        End Property



        Private _sRedirectInfo As String
        ''' <summary>
        ''' Set how the redirect url is being passed or stored.
        ''' Can be stored as cookie or passed as QS parameter.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RedirectInfo() As String
            Get
                Return _sRedirectInfo
            End Get
            Set(ByVal value As String)
                _sRedirectInfo = value
            End Set
        End Property

#End Region


#Region "Legacy Properties"
        'For legacy support and renamed properties

        Public Property IfMobileRX1() As String
            Set(ByVal value As String)
                _sDetectMobileRegex1 = value
            End Set
            Get
                Return _sDetectMobileRegex1
            End Get
        End Property


        Public Property IfMobileRX2() As String
            Set(ByVal value As String)
                _sDetectMobileRegex2 = value
            End Set
            Get
                Return _sDetectMobileRegex2
            End Get
        End Property

#End Region


#Region "Public Readonly Properties"
        ''' <summary>
        ''' Return a string with the browser version
        ''' </summary>
        ''' <param name="ShowMinor"></param>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ShowBrowserVersion(Optional ByVal ShowMinor As Boolean = True) As String
            Get

                Return GetBrowserVersion().Major.ToString


            End Get
        End Property

        ''' <summary>
        ''' Returns the browser name
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ShowBrowser() As String
            Get
                Return GetBrowserName()
            End Get
        End Property


        ''' <summary>
        ''' Return the current culture
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ShowCulture() As String
            Get
                Return CurrentCulture()
            End Get
        End Property

        ''' <summary>
        ''' retrun the current text direction
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ShowIfTextDir() As String
            Get
                If CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft = True Then
                    Return ("rtl")
                Else
                    Return ("ltr")
                End If

            End Get
        End Property

#End Region


#Region "General"




        Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init

            ' If the user has no view rights on this page, redirect them.
            ' This is done here because it must be done before DNN redirects the user
            ' The redirectmode is ignored for this (no cookie can be written)
            If IfNoViewRights Then
                If Not DotNetNuke.Security.Permissions.TabPermissionController.CanViewPage() Then
                    Redirect()
                End If
            End If

            If AddToBodyTop > String.Empty Then
                ProcessAddToBody(AddToBodyTop, InjectPosition.Top)
            End If

            If AddToBodyBottom > String.Empty Then
                ProcessAddToBody(AddToBodyBottom, InjectPosition.Bottom)
            End If


            bConditions = CheckConditions()

            If AddHtmlAttribute <> String.Empty And bConditions Then

                ProcessHtmlAttributes()

            End If

        End Sub


        Private Sub Page_PreRender(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.PreRender


            'Redirect to a page
            'This test is done here to make sure the cookie is written for page the skin object is on
            If bConditions And Not RedirectTo = String.Empty Then
                ProcessRedirect()
            End If

            If Not RemoveFromHead = String.Empty Then
                ProcessRemoveFromHead()
            End If

            If CorrectModuleCss Then
                GoCorrectModuleCss()
            End If

            If ForceSkinCssOrder Then
                GoForceCssOrder()
            End If

            If Not RemoveCssFile = String.Empty Or Not RemoveJsFile = String.Empty Then
                If FilterRemove = False Or bConditions Then
                    ProcessRemoveFiles()
                End If
            End If


            If bConditions Then
                ProcessRemoveControls()
            End If


            If AddLanguageLinksToHead Then


                If FilterAddLanguageLinksToHead = False OrElse bConditions Then

                    Dim strLinks As String = GetLanguageLinks(PortalSettings.ActiveTab.TabID, LanguageLinksTemplate)


                    If strLinks <> String.Empty Then
                        Add2Head(strLinks)
                    End If

                End If


            End If


                If Not RemoveMeta = String.Empty Then
                If FilterRemove = False Or bConditions Then
                    ProcessRemoveMeta()
                End If
            End If


            If Not ChangeMeta = String.Empty Then
                If FilterMeta = False Or bConditions Then
                    ProcessChangeMeta()
                End If
            End If

            If MetaNameToLower Then
                MetaNameLowerCase()
            End If


            If Not AddMetaTags = String.Empty Then
                If FilterMeta = False Or bConditions Then
                    ProcessMetaTags()
                End If
            End If



            If Not (AddCssFile = String.Empty And AddJsFile = String.Empty) Then
                If FilterAdd = False Or bConditions Then
                    ProcessLoadFiles()
                End If
            End If

            If Not AddToHead = String.Empty Then
                If FilterAdd = False Or bConditions Then
                    ProcessAdd2Head()
                End If
            End If




            'Non conditional Options
            If FilterBodyClass = False Or bConditions Then

                If AddBodyClass Then

                    Dim sBC As String = BodyClass

                    If bConditions Then sBC &= " " & AddToBodyClass

                    ProcessPageClassTemplate(sBC)

                End If

            End If


            ' Add the doctype
            ProcessDoctype()


            'Write the text to the HTML attributes literal, the manipulation of it has already been done before this on init.
            Dim oAttributes As Literal = CType(Me.Page.FindControl("attributeList"), Literal)
            If Not oAttributes Is Nothing Then
                oAttributes.Text = HtmlAttributeList()
            End If

            Dim sContent As String = String.Empty


            If bConditions Then
                sContent = Content
            Else
                sContent = ContentFalse
            End If

            ltShowInfo.Text = ProcessTokens(sContent)



            If ShowInfo = True Then WriteShowInfo()

        End Sub




        Private Function AddQueryString(ByVal URL As String, ByVal QS As String) As String
            'Used to add a querystring parameter for the redirect link
            If URL.Contains("?") Then
                Return String.Format("{0}&{1}", URL, QS)
            Else
                Return String.Format("{0}?{1}", URL, QS)
            End If

        End Function






        Private Function GetCpState() As String


            If DotNetNuke.Security.Permissions.TabPermissionController.CanAddContentToPage Or DotNetNuke.Security.Permissions.TabPermissionController.CanManagePage Then
                Select Case PortalSettings.UserMode

                    Case DotNetNuke.Entities.Portals.PortalSettings.Mode.View

                        Return DotNetNuke.Entities.Portals.PortalSettings.Mode.View.ToString

                    Case DotNetNuke.Entities.Portals.PortalSettings.Mode.Edit
                        Return DotNetNuke.Entities.Portals.PortalSettings.Mode.Edit.ToString

                    Case DotNetNuke.Entities.Portals.PortalSettings.Mode.Layout
                        Return DotNetNuke.Entities.Portals.PortalSettings.Mode.Layout.ToString


                End Select

            Else

                If IsModuleEditor() Then
                    Return _strModuleEdit
                End If

            End If

            Return (strUserModeNone)

        End Function





        Private Function CheckCookie(ByVal CookieName As String, ByVal CookieValue As String) As Boolean
            ' Check if the cookie exists, if it has the right value and is not expired..
            ' If CookieValue = "" then retrun True too.

            Dim key As String = HttpContext.Current.Server.UrlEncode(CookieName)

            ' Test if the Cookie exists.
            ' First check in Response
            ' If not found try in the Request object (as otherwise you could get the old and the new value if DNN rewrites the value in this request)

            If HttpContext.Current.Response.Cookies.AllKeys.Contains(key) Then
                'Check if the Cookie value is correct or empty
                If CookieValue = "" OrElse GetResponseCookie(CookieName) = CookieValue Then
                    Return True
                End If

            Else
                ' if not in response, try in request (already on client)
                'Check if the Cookie Exists
                If Not Request.Cookies(CookieName) Is Nothing Then
                    'Check if the Cookie value is correct
                    If CookieValue = "" OrElse Server.HtmlEncode(Request.Cookies(CookieName).Value) = CookieValue Then
                        Return True
                    End If
                End If
            End If


            Return False

        End Function



        Private Function GetResponseCookie(key As String) As String

            'encode key for retrieval
            key = HttpContext.Current.Server.UrlEncode(key)
            Dim ck As HttpCookie = HttpContext.Current.Response.Cookies.Get(key)
            Return (ck.Value)

        End Function




        Private Sub CreateCookie(ByVal CookieName As String, ByVal CookieValue As String, ByVal CookieExpireDays As Integer)

            'Write the cookie
            Response.Cookies(CookieName).Value = CookieValue

            'If on cookie experation passed, don't set it (session cookie)
            If CookieExpireDays > 0 Then
                Response.Cookies(CookieName).Expires = DateTime.Now.AddDays(CookieExpireDays)
            End If

        End Sub



        Private Function CreateCookieName(ByVal Name As String) As String

            Return "40Fingers.StyleHelper." & Name

        End Function




        Private Sub ProcessRemoveFiles()    'Process the files to unload from the head

            If Not RemoveCssFile.Trim = String.Empty Then

                For Each s As String In SplitString(RemoveCssFile, ",")
                    UnloadCss(s.Trim)
                Next

            End If

            If Not RemoveJsFile.Trim = String.Empty Then

                For Each s As String In SplitString(RemoveJsFile, ",")
                    UnloadJs(s.Trim)
                Next

            End If

        End Sub



        Private Sub ProcessRemoveControls()

            If Not RemoveControls.Trim = String.Empty Then

                For Each s As String In SplitString(RemoveControls, SplitChar)
                    RemoveControl(s.Trim)
                Next

            End If

        End Sub


        Private Sub ProcessRemoveMeta()
            'Remove Meta tags, some are literals, some are HtmlMeta controls
            Dim oHead As HtmlGenericControl = CType(Me.Page.FindControl("Head"), HtmlGenericControl)

            For Each s As String In SplitString(RemoveMeta, "||")

                If s.StartsWith("id=") Then 'ID passed, meaning this is a Meta Control
                    s = s.Replace("id=", "")
                    Try
                        Dim oControl As Control = oHead.FindControl(s)
                        If Not oControl Is Nothing Then
                            'Try to parse to meta
                            Dim oMeta As HtmlMeta = CType(oControl, HtmlMeta)
                            oMeta.Visible = False
                        End If
                    Catch ex As Exception
                        DotNetNuke.Services.Exceptions.LogException(ex)
                    End Try
                End If

            Next

        End Sub


        Private Sub ProcessChangeMeta()
            'Remove Meta tags, some are literals, some are HtmlMeta controls
            Dim oHead As HtmlGenericControl = CType(Me.Page.FindControl("Head"), HtmlGenericControl)


            For Each s As String In Regex.Split(ChangeMeta, Regex.Escape("||"))

                'Meta tags are passed like this: id=MyId|content=Test
                If s.StartsWith("id=") Then 'ID passed

                    Dim sAttribute As String = String.Empty ' Store the attribute to look for

                    'Get id and attribute to change if an attribute was passed
                    If s.Split("|"c).Length = 2 Then
                        sAttribute = s.Split("|"c)(1)
                        s = s.Split("|"c)(0)

                        'Get the id of the element to change
                        Dim oId As New ParameterValue(s, "=")
                        s = oId.Value1


                        Dim oValPair As New ParameterValue(sAttribute, "=")

                        Try
                            Dim oControl As Control = oHead.FindControl(s)
                            If Not oControl Is Nothing Then
                                'Try to parse to meta
                                Dim oMeta As HtmlMeta = CType(oControl, HtmlMeta)
                                'Get the parameter
                                If Not oMeta.Attributes(oValPair.Parameter) Is Nothing Then
                                    Dim sNewVal As String = oValPair.Value1
                                    'Replace * with the original value 
                                    sNewVal = sNewVal.Replace("*", oMeta.Attributes(oValPair.Parameter))
                                    'Change the Parameters Value
                                    oMeta.Attributes(oValPair.Parameter) = ProcessTokens(sNewVal)
                                End If

                            End If
                        Catch ex As Exception
                            DotNetNuke.Services.Exceptions.LogException(ex)
                        End Try

                    End If
                End If

            Next

        End Sub



        Private Sub MetaNameLowerCase()

            Dim strAttr As String = "name"
            Dim oHead As HtmlGenericControl = CType(Me.Page.FindControl("Head"), HtmlGenericControl)


            Try
                For Each oControl As Control In oHead.Controls
                    If Not oControl Is Nothing And TypeOf oControl Is HtmlMeta Then
                        'Try to parse to meta
                        Dim oMeta As HtmlMeta = CType(oControl, HtmlMeta)
                        'Get the parameter
                        If Not oMeta.Attributes(strAttr) Is Nothing Then
                            Dim sNewVal As String = oMeta.Attributes(strAttr).ToLower
                            'Replace * with the original value 
                            oMeta.Attributes(strAttr) = sNewVal
                        End If

                    End If

                Next
            Catch ex As Exception
                DotNetNuke.Services.Exceptions.LogException(ex)
            End Try


        End Sub


        Private Sub UnloadCss(ByVal sFileName As String)

            Dim oCSS As Control = Me.Page.FindControl("CSS")

            For Each oControl As Control In oCSS.Controls
                Select Case oControl.GetType.ToString
                    Case "System.Web.UI.HtmlControls.HtmlLink"
                        Dim oLink As HtmlLink = CType(oControl, HtmlLink)
                        If CheckStringFound(oLink.Attributes("href"), sFileName) Then
                            oLink.Visible = False
                        End If
                End Select
            Next

            'For the control panel CSS
            Dim oHead As Control = Me.Page.FindControl("Head")
            For Each oControl As Control In oHead.Controls
                Select Case oControl.GetType.ToString
                    Case "System.Web.UI.LiteralControl"
                        Dim oLink As LiteralControl = CType(oControl, LiteralControl)
                        If CheckStringFound(oLink.Text, sFileName) Then
                            oLink.Visible = False
                        End If
                End Select
            Next

            'For Dnn 6.1+
            Dim oIncludes As Control = Me.Page.FindControl("ClientResourceIncludes")
            If Not oIncludes Is Nothing Then

                'Get list of child items client resource controls
                Dim lstControl2Remove As New List(Of String)


                Dim iItems As Integer = oIncludes.Controls.Count - 1

                'Loop though Items reverse
                For i As Integer = iItems To 0 Step -1

                    Dim oCssControl As Control = oIncludes.Controls(i)

                    'Check if it's a CssInclude
                    Select Case oCssControl.GetType.ToString
                        Case "DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude"
                            Dim oCSSRemove As DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude = CType(oCssControl, DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude)
                            'Check the path
                            If CheckStringFound(oCSSRemove.FilePath, sFileName) Then
                                'Add to the list of controls to remove
                                oIncludes.Controls.RemoveAt(i)

                            End If

                    End Select

                Next


            End If



        End Sub


        ''' <summary>
        ''' Make sure all css is loaded before skin.css, will work only for css injected through Client Resource Management
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub GoCorrectModuleCss()

            'Get the CRM control
            Dim oIncludes As Control = Me.Page.FindControl("ClientResourceIncludes")
            If Not oIncludes Is Nothing Then

                'Get list of child items client resource controls
                Dim CssControls As New List(Of Control)

                'Loop though Items
                For Each oCssControl As Control In oIncludes.Controls()
                    'Check if it's a CssInclude
                    Select Case oCssControl.GetType.ToString
                        Case "DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude"
                            Dim oCssItem As DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude = CType(oCssControl, DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude)
                            'Check the path
                            If Regex.IsMatch(oCssItem.FilePath, "/Desktopmodules/|/admin/|/shared/|/WebControlSkin/", RegexOptions.IgnoreCase) Then
                                'Add to the list of controls to remove
                                If oCssItem.Priority > 15 Then
                                    oCssItem.Priority = 12
                                End If

                            End If

                    End Select
                Next

            End If

        End Sub


        Private Sub GoForceCssOrder()

            'Get the CRM control
            Dim oIncludes As Control = Me.Page.FindControl("ClientResourceIncludes")
            If Not oIncludes Is Nothing Then

                'Get list of child items client resource controls
                Dim CssControls As New List(Of Control)

                'Loop though Items
                For Each oCssControl As Control In oIncludes.Controls()
                    'Check if it's a CssInclude
                    Select Case oCssControl.GetType.ToString
                        Case "DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude"
                            Dim oCssItem As DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude = CType(oCssControl, DotNetNuke.Web.Client.ClientResourceManagement.DnnCssInclude)
                            'Check the path
                            If Regex.IsMatch(oCssItem.FilePath, "/Skin.css|/Portal.css", RegexOptions.IgnoreCase) Then
                                'Add to the list of controls to remove
                                Select Case oCssItem.Priority
                                    Case 15
                                        oCssItem.Priority = 1000
                                    Case 35
                                        oCssItem.Priority = 1001

                                End Select

                            End If

                    End Select
                Next

            End If

        End Sub



        Private Sub UnloadJs(ByVal sFileName As String)


            'For Dnn 6.1+
            Dim oIncludes As Control = Me.Page.FindControl("ClientResourceIncludes")
            If Not oIncludes Is Nothing Then

                'Get list of child items client resource controls
                Dim lstControl2Remove As New List(Of Control)

                'Loop though Items
                For Each oJsControl As Control In oIncludes.Controls()
                    'Check if it's a JsInclude
                    Dim sType As String = oJsControl.GetType.ToString

                    Select Case sType
                        Case "DotNetNuke.Web.Client.ClientResourceManagement.DnnJsInclude"
                            Dim oJsRemove As DotNetNuke.Web.Client.ClientResourceManagement.DnnJsInclude = CType(oJsControl, DotNetNuke.Web.Client.ClientResourceManagement.DnnJsInclude)
                            'Check the path
                            If CheckStringFound(oJsRemove.FilePath, sFileName) Then
                                'Add to the list of controls to remove
                                lstControl2Remove.Add(oJsControl)
                            End If

                    End Select
                Next

                'Loop through found controls and remove them
                For Each oJs2remove As Control In lstControl2Remove

                    oIncludes.Controls.Remove(oJs2remove)

                Next

            End If



        End Sub




        Private Sub ProcessMetaTags()   'Process meta tags to add to the head

            If Not AddMetaTags.Trim = String.Empty Then
                'Get the individual meta Tags
                For Each s As String In SplitString(AddMetaTags, "|")
                    'Split name and Value
                    Dim m As String() = s.Split(":"c)
                    If m.Length = 2 Then
                        WriteMeta(m(0).Trim, m(1).Trim)
                    End If

                Next
            End If

        End Sub



        Public Sub WriteMeta(ByVal sName As String, ByVal sContent As String)

            Dim oMeta As New HtmlMeta
            oMeta.Name = sName
            oMeta.Content = sContent
            Me.Page.Header.Controls.Add(oMeta)


        End Sub



        Private Sub ProcessLoadFiles()  'Process the files to add to the head

            'CSS files
            If Not AddCssFile.Trim = String.Empty Then

                For Each s As String In SplitString(AddCssFile, ",")

                    'First process tokens to prevent issues with css file media passed (uses : too)

                    s = ProcessTokens(s)

                    'Check for "" again as this could be the case if one of the tokens returns nothing

                    If Not s = String.Empty Then

                        Dim sMedia As String = CssMedia

                        'Get the media types, if passed with the filename
                        Dim m As Match = Regex.Match(s, ":[^/].*")

                        If m.Length > 0 Then
                            'Remove the match from the original string
                            s = s.Replace(m.Value, "")
                            'Parse passed tokens

                            'Format the media types
                            sMedia = m.Value
                            sMedia = Regex.Replace(sMedia, "^:", "") 'Remove first :
                            sMedia = Regex.Replace(sMedia, ":", ", ") 'Replace : by ,
                        End If

                        WriteToHead(s, "css", "", AddAtEnd, sMedia)

                    End If

                Next
            End If


            'JS files
            If Not AddJsFile.Trim = String.Empty Then
                For Each s As String In SplitString(AddJsFile, ",")

                    'Process tokens
                    s = ProcessTokens(s)


                    'Check for "" again as this could be the case if one of the tokens returns nothing
                    If Not s = String.Empty Then
                        WriteToHead(s, "js", "", AddAtEnd)
                    End If

                Next
            End If

        End Sub


        Private Sub ProcessAdd2Head()
            'Process the string that is to be added to the head of the page

            Add2Head(AddToHead)

        End Sub


        Private Sub Add2Head(ByVal sIn As String)
            'Write a string to the head of the page
            sIn = PathTokenReplace(sIn)

            'Get the head element
            Dim oHead As HtmlGenericControl = CType(Me.Page.FindControl("Head"), HtmlGenericControl)

            'Create literal and parse the tokens
            Dim litAdd As New Literal
            litAdd.Text = ProcessTokens(sIn)

            'Test if head exists
            If Not oHead Is Nothing Then
                'Get position
                Dim iLocation As Integer = 0
                If AddAtEnd = True Then
                    iLocation = oHead.Controls.Count
                End If
                'Inject the string
                oHead.Controls.AddAt(iLocation, litAdd)

            End If


        End Sub


        Private Sub ProcessRemoveFromHead()
            'Pass attribute and value and test to remove
            'Will not work on most js and css files
            Dim oHead As HtmlGenericControl = CType(Me.Page.FindControl("Head"), HtmlGenericControl)

            For Each s As String In SplitString(RemoveFromHead, "||")

                If Not s = "" Then
                    Dim oValPair As New ParameterValue(s, "=")

                    For Each oControl As Control In oHead.Controls()
                        Select Case oControl.GetType.ToString
                            Case "System.Web.UI.HtmlControls.HtmlLink"

                                'oControl.Visible = False
                                Dim oLink As System.Web.UI.HtmlControls.HtmlLink
                                oLink = CType(oControl, HtmlLink)

                                If oLink.Attributes.Item(oValPair.Parameter) = oValPair.Value1 Then
                                    oControl.Visible = False
                                End If

                            Case "System.Web.UI.HtmlControls.HtmlMeta"

                                'oControl.Visible = False
                                Dim oMeta As System.Web.UI.HtmlControls.HtmlMeta
                                oMeta = CType(oControl, HtmlMeta)
                                If oMeta.Attributes.Item(oValPair.Parameter) = oValPair.Value1 Then
                                    oControl.Visible = False
                                End If

                            Case "System.Web.UI.WebControls.Literal"
                                Dim oLiteral As System.Web.UI.WebControls.Literal
                                oLiteral = CType(oControl, Literal)
                                Dim sFind As String = oValPair.Parameter & "\s?=\s?[""']" & oValPair.Value1 & "[""']"
                                If Regex.IsMatch(oLiteral.Text, sFind, RegexOptions.IgnoreCase) Then
                                    oControl.Visible = False
                                End If

                        End Select


                    Next
                End If

            Next


        End Sub





        Private Sub ProcessHtmlAttributes()
            'Process pipe separated list of attributes for the HTML element
            For Each m As String In SplitString(AddHtmlAttribute, "|")
                Dim s As String() = m.Split(","c)
                If s.Length = 2 Then
                    AddNewHtmlAttribute(m.Split(","c)(0), m.Split(","c)(1))
                End If
            Next

        End Sub



        Private Sub AddNewHtmlAttribute(ByVal Attribute As String, ByVal Value As String)
            'Write attributes to the HTML element
            If Attribute <> String.Empty And Value <> String.Empty Then
                Dim oPage As CDefault = TryCast(Me.Page, CDefault)
                If oPage.HtmlAttributes(Attribute) > "" Then
                    'If attribute already exists, overwrite it
                    oPage.HtmlAttributes(Attribute) = Value
                Else
                    oPage.HtmlAttributes.Add(Attribute, Value)
                End If

            End If

        End Sub


        Protected Function HtmlAttributeList() As String
            'Get the DNN page
            Dim oPage As CDefault = TryCast(Me.Page, CDefault)

            'Are there html attributes?
            If (oPage.HtmlAttributes IsNot Nothing) AndAlso (oPage.HtmlAttributes.Count > 0) Then
                Dim attr As StringBuilder = New StringBuilder("")
                'Loop through the existing attributes
                For Each AttributeName As String In oPage.HtmlAttributes.Keys
                    If (Not [String].IsNullOrEmpty(AttributeName)) AndAlso (oPage.HtmlAttributes(AttributeName) IsNot Nothing) Then
                        Dim attributeValue As String = oPage.HtmlAttributes(AttributeName)
                        If (attributeValue.IndexOf(",") > 0) Then
                            Dim attributeValues() As String = attributeValue.Split(","c)
                            For attributeCounter As Integer = 0 To attributeValues.Length - 1
                                attr.Append((" " & AttributeName & "=""") + attributeValues(attributeCounter) & """")
                            Next
                        Else
                            attr.Append(" " & AttributeName & "=""" & attributeValue & """")
                        End If
                    End If
                Next
                Return attr.ToString()
            End If
            Return ""

        End Function


        Private Sub ProcessDoctype()

            If Doctype <> String.Empty Then
                'Will overwrite the doctype with the passed doctype

                Dim strLang As String = System.Globalization.CultureInfo.CurrentCulture.ToString()
                'Sets the skins doctype
                Dim liDoctype As Literal = CType(Me.Page.FindControl("skinDocType"), Literal)

                If Not liDoctype Is Nothing Then
                    If StringContains(Doctype, "HTML 5|html5") Then 'HTML 5
                        liDoctype.Text = "<!DOCTYPE HTML>"
                    ElseIf StringContains(Doctype, "XHTML") Then 'XHTML

                        AddNewHtmlAttribute("xml:lang", strLang)
                        AddNewHtmlAttribute("xmlns", "http://www.w3.org/1999/xhtml")

                        If StringContains(Doctype, "Strict") Then ' 1.0 Strict
                            liDoctype.Text = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>"
                        ElseIf StringContains(Doctype, "Transitional") Then ' 1.0 Transitional
                            liDoctype.Text = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
                        ElseIf StringContains(Doctype, "XHTML 1.1") Then ' 1.1
                            liDoctype.Text = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.1//EN' 'http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd'>"
                        End If

                    ElseIf StringContains(Doctype, "HTML 4.01 Strict") Then ' HTML 4.01 Strict
                        liDoctype.Text = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/TR/html4/strict.dtd'>"

                    Else
                        'Do nothing..
                    End If
                End If
            End If


        End Sub






        Private Sub WriteToHead(ByVal sFile As String, ByVal sType As String, Optional ByVal sId As String = "", Optional ByVal bAddAtEnd As Boolean = True, Optional ByVal sMedia As String = "screen")

            Dim sFilePath As String = ProcessPath(BasePath, sFile)
            Dim oLink As New HtmlGenericControl

            Select Case sType 'Check what kind of file this is

                Case "css" ' CSS
                    oLink = New HtmlGenericControl("link")
                    oLink.Attributes("href") = sFilePath
                    oLink.Attributes("media") = sMedia
                    oLink.Attributes("type") = "text/css"
                    oLink.Attributes("rel") = "stylesheet"
                Case "js" ' JS
                    oLink = New HtmlGenericControl("script")
                    oLink.Attributes("language") = "javascript"
                    oLink.Attributes("type") = "text/javascript"
                    oLink.Attributes("src") = sFilePath
            End Select

            If Not oLink Is Nothing Then
                Dim oAddTo As Control


                Dim iLocation As Integer = 0
                If bAddAtEnd = True Then
                    oAddTo = Me.Page.FindControl("Head")
                    iLocation = oAddTo.Controls.Count
                Else
                    oAddTo = Me.Page.FindControl("CSS")
                End If

                If Not oAddTo Is Nothing Then
                    oAddTo.Controls.AddAt(iLocation, oLink)
                End If
            End If

        End Sub



        Private Sub RemoveControl(ByVal id As String)
            'Set a controls visibilty to false
            Dim oForm As Control = Me.Page.FindControl("Form")

            DisableControlRecursive(oForm, id)

        End Sub



        Public Sub DisableControlRecursive(C As Control, ID As String)
            ' Recursively look for control to remove

            For Each childControl As Control In C.Controls

                Dim sID As String = childControl.ID

                If Not sID Is Nothing Then
                    If sID.ToLower = ID.ToLower Then
                        childControl.Visible = False
                        Exit Sub
                    Else
                        DisableControlRecursive(childControl, ID)
                    End If
                Else
                    DisableControlRecursive(childControl, ID)
                End If

            Next
        End Sub



        Private Function ProcessPath(ByVal BasePath As String, ByVal Path As String) As String

            'Check if the subpath is not a absolute URI and there is no root token in the base path.
            If Not (Regex.IsMatch(Path, "\[.\]", RegexOptions.IgnoreCase) OrElse Path.Contains("://") OrElse Path.StartsWith("/")) Then
                Path = CombinePath(PathTokenReplace(BasePath), Path)
            Else
                Path = PathTokenReplace(Path)
            End If

            Return Path

        End Function


        Private Function PathTokenReplace(ByVal Path As String) As String

            If Path.Contains("[") Then

                'Replace tokens in Path
                Path = Regex.Replace(Path, "\[S\]", PortalSettings.ActiveTab.SkinPath, RegexOptions.IgnoreCase) 'Skin Path
                Path = Regex.Replace(Path, "\[P\]", PortalSettings.HomeDirectory, RegexOptions.IgnoreCase) 'Portal Path
                Path = Regex.Replace(Path, "\[M\]", Me.TemplateSourceDirectory, RegexOptions.IgnoreCase) 'Skin Object Path
                Path = Regex.Replace(Path, "\[R\]", "/") 'Root Path
                Path = Regex.Replace(Path, "\[D\]", "/DesktopModules/") 'DesktopModules

            End If

            Return Path

        End Function


        Private Function CombinePath(ByVal BasePath As String, ByVal SubPath As String) As String
            'Combine two path

            If BasePath.EndsWith("/") Then
                If SubPath.StartsWith("/") Then
                    SubPath.Substring(1, SubPath.Length - 1)
                End If
            Else
                If Not SubPath.StartsWith("/") Then
                    SubPath = "/" & SubPath
                End If
            End If

            Return BasePath & SubPath


        End Function



        Private Function CheckConditions() As Boolean

            Return CheckBrowsers(IfBrowser) AndAlso _
            CheckDnnVersion(IfDnnVersion) AndAlso _
            CheckUser(IfUserName) AndAlso _
            CheckUserMode(IfUserMode) AndAlso _
            CheckURLs(IfURL) AndAlso _
            CheckCultures(IfCulture) AndAlso _
            CheckUserAgents(IfUserAgentString) AndAlso _
            CheckRoles(IfRole) AndAlso _
            CheckQueryString(IfQS) AndAlso _
            CheckTextDir(IfTextDir) AndAlso _
            CheckMobile() AndAlso _
            CheckNoCookies(IfNoCookie) AndAlso _
            CheckCookies(IfCookie) AndAlso _
            CheckTokens(IfToken) AndAlso _
            _bIfExternal = True


        End Function


#End Region



#Region "Redirects"



        Private Sub ProcessRedirect()       'Process redirect

            'Check first if the redirection should be stopped.
            Dim iStopRedirect As Integer = StopRedirect()

            Dim sCookieName As String = CreateCookieName(String.Format("RedirectStop.{0}.{1}.{2}", RedirectStop, RedirectMode, RedirectName))


            Select Case RedirectMode.ToLower



                Case "never" 'Don't redirect, allows disabeling redirect for testing..
                    Exit Sub

                Case "once" 'Redirect once (cookie valid for one year)

                    'Check if the redirect was refused before
                    If Not CheckCookie(sCookieName, "True") Then
                        Select Case iStopRedirect
                            Case 0
                                Redirect()
                            Case 1
                                CreateCookie(sCookieName, "True", 365)
                            Case 2
                                CreateCookie(sCookieName, "True", 365)
                                Redirect()
                        End Select
                    End If

                Case "session", "oncepersession" 'Redirect once this session but also redirect the next session..

                    'Check if the redirect was refused before
                    If Not CheckCookie(sCookieName, "True") Then
                        Select Case iStopRedirect
                            Case 0
                                Redirect()
                            Case 1
                                CreateCookie(sCookieName, "True", 0)
                            Case 2
                                CreateCookie(sCookieName, "True", 0)
                                Redirect()
                        End Select

                    End If

                Case "always"
                    If iStopRedirect = 0 Then
                        Redirect()
                    End If
            End Select

        End Sub


        Private Function StopRedirect() As Integer
            'Check if you should stop redirecting (set with the RedirectStop attribute)
            '0 = don't stop
            '1 = Stop now
            '2 = Stop next time

            If RedirectStop.ToLower.Contains("never") Then Return 0

            'Is this page was redirected before, redirect will find out if this is a revisit
            If RedirectStop.ToLower.Contains("revisit") Then Return 2

            'If a querystring parameter was passed
            If RedirectStop.ToLower.Contains("querystring") Then
                'Check for QS parameter
                If CheckQueryString("NoRedirect:True") Then
                    Return 1
                End If
            End If

            'If the referer is the redirect url 
            If RedirectStop.ToLower.Contains("redirecturl") Then

                'Get the Referrer URL
                Dim sReferrer As String = String.Empty
                If Not Request.UrlReferrer Is Nothing Then 'Check if there was a referrer
                    sReferrer = Request.UrlReferrer.GetLeftPart(UriPartial.Path)
                End If

                If RedirectStop.ToLower.Contains("baseredirecturl") Then
                    'Check the referer against the passed referer Base URL (passed in Skin)
                    If RedirectBaseUrl > String.Empty AndAlso sReferrer.StartsWith(RedirectBaseUrl) Then
                        Return 1
                    End If
                Else
                    'Test if the referring page is the page to redirect to
                    If sReferrer = RedirectTo Or sReferrer.EndsWith(RedirectTo) Then
                        Return 1
                    End If
                End If

            End If

            Return 0


        End Function



        Private Sub Redirect()
            'Do the actual redirect
            Dim RedirectURL As String = RedirectTo

            If RedirectInfo <> String.Empty Then
                'Writes cookie or querystring with the value of the original page url
                Dim BaseName As String = "RedirectedFrom"

                Dim sOrignalURL As String = PortalSettings.ActiveTab.FullUrl

                'Write a cookie with the original landing page
                If Regex.IsMatch(RedirectInfo, "Cookie", RegexOptions.IgnoreCase) Then 'Write cookie
                    Dim CookieName As String = CreateCookieName(BaseName)
                    Response.Cookies(CookieName).Value = sOrignalURL
                End If

                'Add redirect ot QueryString
                If Regex.IsMatch(RedirectInfo, "QS", RegexOptions.IgnoreCase) Then 'Add to querystring
                    RedirectURL = AddQueryString(RedirectTo, String.Format("{0}={1}", BaseName, HttpUtility.UrlEncode(sOrignalURL)))
                End If

            End If

            Response.Redirect(GetRedirectUrl(RedirectURL))

        End Sub



        Private Function GetRedirectUrl(ByVal URL As String) As String
            'Input: the new URL template
            'Example: m.test.com?[QS:DetailId]

            'Get the complete URL
            URL = ProcessPath(RedirectBaseUrl, URL)

            'Get the passed tokens in the new URL
            Dim sUrlTemp As String = URL

            For Each m As Match In Regex.Matches(sUrlTemp, "\[(.*)\]")

                'Remove found token
                URL = Regex.Replace(URL, Regex.Escape(m.Value), "")

                'Get parameter and value to test for
                Dim myPV As New ParameterValue(m.Groups(1).Value, ":")

                Dim sQsPm As String = String.Empty

                If myPV.Parameter = "QS" Then
                    'Get the QS parameter from the original URL
                    sQsPm = GetQsValue(myPV.Value1)
                End If

                'Test if the value was found...
                If Not sQsPm = String.Empty Then
                    'Add qs parameter
                    URL = AddQueryString(URL, myPV.Value1 & "=" & sQsPm)
                End If

            Next

            Return (URL)

        End Function



#End Region


#Region "Condition Functions"


        Protected Function CheckDnnVersion(condition As String) As Boolean

            If condition = String.Empty Then
                Return True
            Else


                Return CheckVersion(condition, DotNetNukeContext.Current.Application.Version)

            End If


        End Function


        Protected Function CurrentDnnVersion(Digits As Integer) As String

            Dim oVersion As Version = DotNetNukeContext.Current.Application.Version

            Dim strFormat As String = "{0:00}"

            Dim sOut As String = String.Format(strFormat, oVersion.Major)

            If Digits >= 2 Then

                sOut &= String.Format(strFormat, oVersion.Minor)

            End If

            If Digits >= 3 Then

                sOut &= String.Format(strFormat, oVersion.MinorRevision)

            End If


            Return (sOut)


        End Function

        Protected Function CheckBrowsers(ByVal sBrowsers As String) As Boolean
            'Check if the current browser is part of the browser string

            If sBrowsers = String.Empty Then
                Return (True)
            Else
                Dim bOut As Boolean = False
                For Each s As String In SplitString(sBrowsers, ",")
                    bOut = (CheckBrowser(s))
                    If bOut Then
                        Return (bOut)
                    End If
                Next
                Return (bOut)
            End If

        End Function


        ''' <summary>
        ''' Check if this is the passed browser and if it's the passed version number
        ''' </summary>
        ''' <param name="sCheck"></param>
        ''' <returns></returns>

        Protected Function CheckBrowser(ByVal sCheck As String) As Boolean

            Dim bReturn As Boolean = False

            sCheck = sCheck.ToLower


            'Get all the current browsers name and version 
            Dim sBrowserName As String = GetBrowserName()
            Dim oBrowserVersion As Version = GetBrowserVersion()

            'Check if the result should be "inverted" i.e check starts with a !
            Dim bInvert As Boolean = False


            'Check for Negation (!, this will be removed from sCheck)
            bInvert = Not ParseInvert(sCheck)


            'Split Browser name and Version
            'location of the split character
            Dim iSplit As Integer = sCheck.IndexOfAny(caCondition)

            'Individual comparison strings
            Dim sCheckBrowserName As String = String.Empty
            Dim sCheckVersion As String = String.Empty


            'Get the Browsers and version condition if a condition exists
            If iSplit > -1 Then

                sCheckBrowserName = sCheck.Substring(0, iSplit)
                sCheckVersion = sCheck.Substring(iSplit)

            Else

                sCheckBrowserName = sCheck

            End If


            'If the browser name matches
            If (sCheckBrowserName.Trim = sBrowserName) Then

                'If there is no condition
                If sCheckVersion = "" Then
                    'If there is no version number
                    bReturn = True
                Else

                    bReturn = CheckVersion(sCheckVersion, oBrowserVersion)

                End If
            End If

            Return (Not (bReturn Xor bInvert))

        End Function


        Protected Function CheckVersion(check As String, CurrentVersion As String) As Boolean

            Dim ver As New Version(CurrentVersion)

            Return (CheckVersion(check, ver))


        End Function


        Protected Function CheckVersion(check As String, CurrentVersion As Version) As Boolean
            'pass a check, "=10.2", ">10.2", "<10.2" or ">=10", compare with the passed version and return true or false
            Dim bValue As Boolean = False

            Dim charDigits As Char = "."c

            'Split condition and Version
            'location of the split character
            Dim iSplit As Integer = check.LastIndexOfAny(caCondition)

            'Individual comparison strings
            Dim sCheckCondition As String = String.Empty
            Dim sCheckVersion As String = String.Empty


            'Get the if a condition exists
            If iSplit > -1 Then

                sCheckCondition = check.Substring(0, iSplit + 1)
                sCheckVersion = check.Substring(iSplit + 1)

            Else
                ' if not use equal
                sCheckCondition = "="
                sCheckVersion = check


            End If


            Dim iDigits As Integer = sCheckVersion.Split(charDigits).Length


            If iDigits < 2 Then

                sCheckVersion = String.Format("{0}.00", sCheckVersion)

            End If

            Dim verCheck As Version


            Try

                verCheck = New Version(sCheckVersion)

            Catch ex As Exception

                verCheck = New Version("99.99")

            End Try


            Dim result As VersionIs = CompareVersions(CurrentVersion, verCheck, iDigits)


            'In case there was no comparison character passed



            'Get the kind of filtering
            If sCheckCondition.Contains("=") Then
                If result = VersionIs.Same Then bValue = True
            End If

            If sCheckCondition.Contains("<") Then
                If result = VersionIs.Smaller Then bValue = True
            End If

            If sCheckCondition.Contains(">") Then
                If result = VersionIs.Bigger Then bValue = True
            End If


            Return bValue


        End Function




        Function CompareVersions(ByVal Version As Version, ByVal CompVersion As Version, ByVal CompareParts As Integer) As VersionIs

            ' Pass two net version objects and compare them
            ' CompareParts, sets what part of the version number should be compared (00.00.00) = 3, (00.00)= 2


            If Version Is Nothing Then

                Return VersionIs.Smaller

            End If

            If CompVersion Is Nothing Then

                Return VersionIs.Bigger

            End If

            ' Only make the comparison on bigger or smaller for each relevant part
            ' Same will then be the end result if nothing matched

            If CompareParts >= 1 AndAlso Version.Major <> CompVersion.Major Then

                If Version.Major > CompVersion.Major Then
                    Return VersionIs.Bigger
                Else
                    Return VersionIs.Smaller
                End If

            End If

            If CompareParts >= 2 AndAlso Version.Minor <> CompVersion.Minor Then

                If Version.Minor > CompVersion.Minor Then
                    Return VersionIs.Bigger
                Else
                    Return VersionIs.Smaller
                End If

            End If

            If CompareParts >= 3 AndAlso Version.Build <> CompVersion.Build Then

                If Version.Build > CompVersion.Build Then
                    Return VersionIs.Bigger
                Else
                    Return VersionIs.Smaller
                End If

            End If

            If CompareParts >= 4 AndAlso Version.Revision <> CompVersion.Revision Then

                If Version.Revision > CompVersion.Revision Then
                    Return VersionIs.Bigger
                Else
                    Return VersionIs.Smaller
                End If

            End If

            'No smaller or bigger matched so this must be equal

            Return VersionIs.Same


        End Function





        Protected Function GetBrowserVersion() As Version

            'Get a version object for the users browser

            ' As there are some robots that pass very weird version numbers, the safest way around this is a try catch.
            Dim strDefaultVersion As String = "0.0.0"
            Try


                Dim sVersion As String

                If Not Request.Browser.Version Is Nothing Then
                    sVersion = Request.Browser.Version.Replace(",", ".")
                Else
                    sVersion = strDefaultVersion
                End If

                Return New Version(sVersion)

            Catch

                Return New Version(strDefaultVersion)

            End Try



        End Function




        Private Function CheckUserAgents(ByVal sUserAgents As String) As Boolean

            If sUserAgents = String.Empty Then
                Return (True)
            Else
                'Get the current UserAgentString
                Dim sUserAgent As String = Request.UserAgent
                Return (MatchString(sUserAgent, sUserAgents))
            End If

        End Function




        Private Function CheckUser(ByVal sUsers As String) As Boolean

            If sUsers = String.Empty Then
                Return (True)
            Else
                'Get the current username
                Dim sUser As String = UserController.Instance.GetCurrentUserInfo().Username
                Return (MatchString(sUser, sUsers))
            End If

        End Function


        Private Function CheckUserMode(ByVal sUserModes As String) As Boolean


            If sUserModes = String.Empty Then
                Return (True)
            Else
                'Get the current usermode
                Dim sUserMode As String = GetCpState()

                Return (MatchString(sUserMode, sUserModes))

            End If

        End Function




        Private Function CheckRoles(ByVal sRoles As String) As Boolean


            If sRoles = String.Empty Then
                'If no filter roles set return True
                Return (True)
            Else



                'Regular Check, user in roles
                Dim sAllroles As String = GetUserroles()

                If sRoles.ToLower = "none" Then
                    If sAllroles = String.Empty Then
                        Return True
                    Else
                        Return False
                    End If
                End If


                'As there is no superuser role, so add it for superusers
                If UserController.Instance.GetCurrentUserInfo().IsSuperUser Then
                    sAllroles = CharSepStrAdd(sAllroles, "SuperUsers", "|")
                End If


                'If the user is not a member of any roles (not logged in)
                If sAllroles = "" Then
                    If (MatchString("*", sRoles)) Then
                        Return True
                    End If
                Else

                    'If is a member of a roles
                    If (MatchString(sAllroles, sRoles)) Then
                        Return True
                    End If

                    'If SuperUsers was part of the roles list

                End If

                Return False

            End If

        End Function




        Private Function GetUserroles() As String
            'Gets the roles a user is in as one string for regex comparison
            Dim sOut As String = String.Empty

            'Create string with all roles for this user in it
            For Each sRole As String In UserController.Instance.GetCurrentUserInfo().Roles
                sOut = CharSepStrAdd(sOut, sRole, "|")
            Next

            Return sOut

        End Function


        Private Function CharSepStrAdd(ByVal Original As String, ByVal Add As String, ByVal Separator As String) As String
            'Add a character separated string to an existing string

            If Add > String.Empty Then
                If Original = String.Empty Then
                    Return Add
                Else
                    Return String.Format("{0}{1}{2}", Original, Separator, Add)
                End If
            Else
                Return Original
            End If

        End Function

        Private Function CheckRoles2(ByVal sRoles As String) As Boolean


            If sRoles = String.Empty Then
                'If no filter roles set return True
                Return (True)
            Else
                'If the user is not a member of any roles
                If UserController.Instance.GetCurrentUserInfo().Roles.Length = 0 Then
                    If (PortalSettings.Current.UserInfo.IsSuperUser) Then
                        If (MatchString(PortalSettings.AdministratorRoleName, sRoles)) Then
                            Return True
                        End If
                    Else
                        If (MatchString("*", sRoles)) Then
                            Return True
                        End If
                    End If
                Else
                    'Regular Check, user in roles
                    Dim sAllroles As String = String.Empty
                    'Create string with all roles for this user in it
                    For Each sRole As String In UserController.Instance.GetCurrentUserInfo().Roles
                        sAllroles &= sRole
                    Next
                    If (MatchString(sAllroles, sRoles)) Then
                        Return True
                    End If

                    'If SuperUsers was part of the roles list

                    If UserController.Instance.GetCurrentUserInfo().IsSuperUser Then
                        If (MatchString(sRoles, "superuser")) Then
                            Return True
                        End If
                    End If


                End If

                Return False

            End If


        End Function




        Private Function CheckQueryString(ByVal sQS As String) As Boolean
            'Check if the query string parameter exists and if the value is correct

            'If no QS passed, skip check
            If sQS = String.Empty Then
                Return (True)
            Else


                'Get parameter and value to test for
                Dim ifQS As New ParameterValue(sQS, ":")

                Dim QSValue As String = GetQsValue(ifQS.Parameter)


                If ifQS.Value1 = "" Then
                    'This means only a QS parameter was passed a condition, return true if it exists in the current URL
                    If QSValue > "" Then
                        Return True
                    End If
                Else
                    'Check the passed QS paramter value against the current URL's QS value

                    If (MatchString(QSValue.ToLower, ifQS.Value1.ToLower)) Then
                        Return (True)
                    End If

                End If

                Return (False)

            End If

        End Function




        Private Function CheckURLs(ByVal sUrls As String) As Boolean

            If sUrls = String.Empty Then
                Return True
            Else
                'Get the current URL
                Dim sUrl As String = GetFullUrl()


                Return (MatchString(sUrl, sUrls))

            End If

        End Function




        Protected Function CheckCultures(ByVal sCultures As String) As Boolean


            If sCultures > String.Empty Then
                'Get the current username
                Dim sCulture As String = CurrentCulture.ToLower

                Return (MatchString(sCulture, sCultures))
            Else
                Return (True)

            End If

        End Function




        Protected Function CheckTextDir(ByVal sTextDir As String) As Boolean
            'Check if there should be a file loaded for RTL or LTR

            If Not sTextDir = String.Empty Then

                Return Not ((sTextDir = "rtl") Xor CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft)

            End If

            Return (True)

        End Function


        Protected Function CheckCookies(ByVal sCookies As String) As Boolean
            'Check if the passed cookies

            If sCookies = String.Empty Then
                Return (True)
            Else
                Dim bOut As Boolean = True
                For Each s As String In SplitString(sCookies, "||")
                    Dim ParVal As New ParameterValue(s, ":")
                    If Not CheckCookie(ParVal.Parameter, ParVal.Value1) Then
                        Return False
                    End If
                Next
                Return (bOut)
            End If

        End Function



        Protected Function CheckNoCookies(ByVal sNoCookies As String) As Boolean
            'Check if the passed cookies don't exist

            If sNoCookies = String.Empty Then
                Return (True)
            Else
                Dim bOut As Boolean = True
                For Each s As String In SplitString(sNoCookies, "||")

                    If CheckCookie(s, "") Then
                        Return False
                    End If

                Next
                Return (bOut)
            End If

        End Function


        ''' <summary>
        ''' Test to see if the parsed token striong returns anything.
        ''' If it does return True
        ''' If it doesn't return False
        ''' </summary>
        ''' <param name="sTokens"></param>
        ''' <returns></returns>
        Protected Function CheckTokens(sTokens As String) As Boolean

            If sTokens = String.Empty Then
                Return True
            Else


                If ProcessTokens(sTokens).Trim = String.Empty Then
                    Return False
                Else
                    Return True
                End If

            End If

        End Function


#End Region


#Region "ShowInfo"

        Private Sub WriteShowInfo()
            'This will write the detected information to the page for testing
            Dim s As String = "<div class=""ff_sh_item"">{0}: {1}</div>"
            Dim sbOut As New StringBuilder
            sbOut.AppendLine("<div class=""FFStyleHelper""><h4>40FINGERS Style Helper Skin Object</h4>")
            sbOut.AppendLine(String.Format(s, "ASP.NET detected Browser", ShowBrowser))
            sbOut.AppendLine(String.Format(s, "ASP.NET detected Browserversion", ShowBrowserVersion))
            sbOut.AppendLine(String.Format(s, "Useragent String", Request.UserAgent))
            sbOut.AppendLine(String.Format(s, "Current Username", UserController.Instance.GetCurrentUserInfo().Username))

            Dim sRoles As String = String.Empty

            For Each sRole As String In UserController.Instance.GetCurrentUserInfo().Roles
                If Not sRole = String.Empty Then
                    sRole &= ", "
                End If
                sRoles &= sRole
            Next

            sbOut.AppendLine(String.Format(s, "Current UserRoles", sRoles))
            sbOut.AppendLine(String.Format(s, "Culture", ShowCulture))
            sbOut.AppendLine(String.Format(s, "Text Direction", ShowIfTextDir))
            sbOut.AppendLine("</div>")

            ltShowInfo.Text &= sbOut.ToString


        End Sub

#End Region


#Region "Helper Functions"



        Protected Function GetBrowserName() As String

            Dim sOut As String = "Undefined"


            If Not Request.Browser.Browser Is Nothing Then


                sOut = Request.Browser.Browser.ToLower

                'Correct change of returned browser for IE 11
                If sOut = "internetexplorer" Then sOut = "ie"
                If sOut = "internet explorer" Then sOut = "ie"

                'Workaround for Edge (not in the browser capabilites file)
                If Not Request.UserAgent Is Nothing AndAlso Regex.IsMatch(Request.UserAgent, "Edge\/\d+") Then
                    sOut = "edge"
                End If

            End If




            Return sOut

        End Function




        Private Function CurrentCulture() As String

            Return CultureInfo.CurrentUICulture.Name.ToString

        End Function


        Private Function CurrentLanguage() As String

            Return CultureInfo.CurrentUICulture.Name.ToString.Substring(0, 2)

        End Function



        Private Function SplitString(Base As String, Split As String) As Array

            ' Splits a string in an array of string, based on a SplitString
            Dim sSplit As String = Regex.Escape(Split)

            Return Regex.Split(Base, sSplit, RegexOptions.IgnoreCase)


        End Function


        Private Function IsModuleEditor() As Boolean

            Dim _IsModuleAdmin As Boolean = Null.NullBoolean

            '' Loop over all modules on this page
            For Each objModule As ModuleInfo In TabController.CurrentPage.Modules

                If Not objModule.IsDeleted Then
                    Dim blnHasModuleEditPermissions As Boolean = ModulePermissionController.HasModuleAccess(SecurityAccessLevel.Edit, Null.NullString, objModule)

                    If blnHasModuleEditPermissions Then
                        _IsModuleAdmin = True
                        Exit For
                    End If
                End If

            Next

            Return PortalController.Instance.GetCurrentPortalSettings().ControlPanelSecurity = PortalSettings.ControlPanelPermission.ModuleEditor AndAlso _IsModuleAdmin


        End Function











        Private Class ParameterValue
            'Pass a QS parameter & value combination and split in parameter and value

            Private _sParameter As String
            Public Property Parameter() As String
                Get
                    Return _sParameter
                End Get
                Set(ByVal value As String)
                    _sParameter = value
                End Set
            End Property



            Private _sValue As String
            Public Property Value1() As String
                Get
                    Return _sValue
                End Get
                Set(ByVal value As String)
                    _sValue = value
                End Set
            End Property



            Public Sub New(ByVal Pair As String, ByVal SplitChar As String)

                SplitChar = Regex.Escape(SplitChar)

                If Pair > "" Then
                    'Check if any value was passed
                    If Pair.Contains(SplitChar) Then
                        Parameter = Regex.Split(Pair, SplitChar, RegexOptions.IgnoreCase)(0)
                        Value1 = Regex.Split(Pair, SplitChar, RegexOptions.IgnoreCase)(1)
                    Else
                        'If no value passed, return "" as the value
                        Parameter = Pair
                        Value1 = ""
                    End If

                End If

            End Sub


        End Class

        Public Function GetValue(ByVal Pair As String, ByVal SplitChar As String) As String

            If Pair > "" Then
                Dim oPV As New ParameterValue(Pair, SplitChar)
                Return oPV.Value1
            End If

            Return ""

        End Function



        Private Function GetQsValue(ByVal Parameter As String) As String

            If Request.QueryString(Parameter) Is Nothing Then
                Return (String.Empty)
            Else
                Return Request.QueryString(Parameter)
            End If

        End Function


        Private Function RemoveStartEnd(ByVal sIn As String, ByVal StartString As String, ByVal EndString As String) As String
            'Remove characters a start eand end of string

            StartString = "^" & Regex.Escape(StartString)
            sIn = Regex.Replace(sIn, StartString, "", RegexOptions.IgnoreCase)

            EndString = Regex.Escape(EndString) & "$"
            sIn = Regex.Replace(sIn, EndString, "", RegexOptions.IgnoreCase)

            Return sIn

        End Function




        Private Function MatchString(ByVal sTest As String, ByVal sParameters As String) As Boolean
            'Test if one of the parameters (comma separated) matches the Test string
			' This functions as an OR for comma separated values, as the function is exited as soon as a value is found.

            'Noting passed
            If sParameters = String.Empty Or sTest = String.Empty Then
                Return False
            End If

            Dim bReturn As Boolean = False




            'Split the comma separated list and parse them one by one.
            For Each s As String In Regex.Split(sParameters, ",", RegexOptions.IgnoreCase)

                'Check if the string contains an invert
                Dim bInvert As Boolean = ParseInvert(s)

                'Escape the string for regex
                s = Regex.Escape(s)

                'Check if the passed value is found
                Dim bFound As Boolean = Regex.IsMatch(sTest, s, RegexOptions.IgnoreCase)

                'If found value needs to be inverted
                If bInvert Then
                    If bFound = True Then
                        Return (False)
                    Else
                        'This can be the correct value, but the rest of the test needs bo be done (for !a, !b)
                        bReturn = True
                    End If
                Else
                    'For a normal value, if the value is found this must be correct, if not it must be incorrect
                    If bFound = True Then
                        Return True
                    Else
                        bReturn = False
                    End If

                End If

            Next

            Return bReturn

        End Function




        Private Function CheckStringFound(ByVal sIn As String, ByVal sCheck As String) As Boolean

            Try
                If sCheck = "/" Then sCheck = "/.*"
                Return Regex.IsMatch(sIn, sCheck, RegexOptions.IgnoreCase)
            Catch ex As Exception
                Return False
            End Try

        End Function




        Private Function ParseInvert(ByRef sIn As String) As Boolean
            'Check for Invert (!) and return the passed value without it
            Dim sTest As String = "^" & _sNOT

            If Regex.IsMatch(sIn, sTest) Then
                sIn = Regex.Replace(sIn, sTest, "")
                Return True
            Else
                Return False
            End If

        End Function



        Private Function StringContains(ByVal Input As String, ByVal Contains As String) As Boolean

            If Regex.IsMatch(Input, Contains, RegexOptions.IgnoreCase) Then
                Return True
            Else
                Return False
            End If

        End Function



        Private Function ProcessTokens(ByVal sOriginal As String) As String


            Dim sResult As String = ProcessTokens(sOriginal, TextType.Text)

            Return sResult

        End Function



        Private Function ProcessTokens(ByVal sOriginal As String, Type As TextType) As String


            If sOriginal.Contains("[") Then


                'Portal
                Dim strPortal As String = "[Portal"

                'Is there a portal token
                If sOriginal.Contains(strPortal) Then

                    Dim strProtocol As String = GetAliasProtocol()
                    Dim strParentAlias As String = GetParentAlias()



                    'Replace PortalId
                    sOriginal = ReplaceToken(sOriginal, "[Portal:Id]", PortalSettings.PortalId.ToString, TextType.Text)
                    sOriginal = ReplaceToken(sOriginal, "[PortalId]", PortalSettings.PortalId.ToString, TextType.Text) ' Legacy

                    sOriginal = ReplaceToken(sOriginal, "[Portal:Alias]", PortalSettings.PortalAlias.HTTPAlias, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Portal:Alias.Root]", strParentAlias, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Portal:Alias.Protocol]", strProtocol, Type)

                    sOriginal = ReplaceToken(sOriginal, "[Portal:Name]", PortalSettings.PortalName, Type)

                    sOriginal = ReplaceToken(sOriginal, "[Portal:Logo]", PortalSettings.LogoFile, Type)

                    sOriginal = ReplaceToken(sOriginal, "[Portal:Logo.Path]", strProtocol & "://" & strParentAlias & PortalSettings.HomeDirectory & PortalSettings.LogoFile, Type)

                    sOriginal = ReplaceSettingsToken(sOriginal, SettingsType.Portal, Type)

                End If

                'Pages / Tabs

                'Portal
                Dim strPage As String = "[Page"

                'Is there a portal token
                If sOriginal.Contains(strPage) Then

                    'Replace Page data
                    Dim oTab As DotNetNuke.Entities.Tabs.TabInfo = GetTabData()

                    sOriginal = ReplaceToken(sOriginal, "[Page:Name]", oTab.TabName, Type)
                    sOriginal = ReplaceToken(sOriginal, "[PageName]", oTab.TabName, Type) ' Legacy

                    sOriginal = ReplaceToken(sOriginal, "[Page:Level]", oTab.Level.ToString, TextType.Text)
                    sOriginal = ReplaceToken(sOriginal, "[PageLevel]", oTab.Level.ToString, TextType.Text) ' Legacy

                    sOriginal = ReplaceToken(sOriginal, "[Page:Title]", oTab.Title, Type)

                    sOriginal = ReplaceToken(sOriginal, "[Page:Description]", oTab.Description, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Page:Url]", GetFullUrl, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Page:RelativeUrl]", System.Web.HttpContext.Current.Request.RawUrl, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Page:Id]", oTab.TabID.ToString, TextType.Text)

                    sOriginal = ReplaceToken(sOriginal, "[Page:Skin]", oTab.SkinSrc, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Page:Container]", oTab.ContainerSrc, Type)

                    sOriginal = ReplaceToken(sOriginal, "[Page:IconFile]", oTab.IconFile, Type)
                    sOriginal = ReplaceToken(sOriginal, "[Page:IconFileLarge]", oTab.IconFileLarge, Type)


                    sOriginal = ReplaceSettingsToken(sOriginal, SettingsType.Tab, Type)


                End If


                sOriginal = ReplaceToken(sOriginal, "[Date]", Now().ToString("yyyyMMdd"), TextType.Text)

                'As these are calculated Values, first check if there is a match for the Token

                'Get Control Panel Class
                If MatchToken(sOriginal, "[CPState]") Then
                    sOriginal = ReplaceToken(sOriginal, "[CPState]", GetCpState, Type)
                End If

                'Get DNN Version Class, but only for unauthenticated users

                If MatchToken(sOriginal, "[DnnVersion]") Then

                    Dim val As String = String.Empty

                    If Not GetCpState() = strUserModeNone Then

                        val = "DNN" & CurrentDnnVersion(1)

                    End If

                    sOriginal = ReplaceToken(sOriginal, "[DnnVersion]", val, TextType.CssClass)

                End If


                'Get Culture Class
                If MatchToken(sOriginal, "[Culture]") Then
                    sOriginal = ReplaceToken(sOriginal, "[Culture]", CurrentCulture, Type)
                End If


                'Get Language Class
                If MatchToken(sOriginal, "[Language]") Then
                    sOriginal = ReplaceToken(sOriginal, "[Language]", CurrentLanguage, Type)
                End If


                'Get PageType Class
                If MatchToken(sOriginal, "[PageType]") Then
                    sOriginal = ReplaceToken(sOriginal, "[PageType]", GetPageType, TextType.Text)
                End If


                'Get Browser Class
                If MatchToken(sOriginal, "[IE]") Then
                    sOriginal = ReplaceToken(sOriginal, "[IE]", GetBrowserClass(12, "ie"), TextType.Text)
                End If


                ' Still tokens?
                If sOriginal.Contains("[") Then

                    'Replace QS Parameters

                    'Get the pattern [xxx]
                    Dim m As Match = Regex.Match(sOriginal, "\[.*:.*\]", RegexOptions.IgnoreCase)

                    If Not m.Value = String.Empty Then

                        'Get rid of the []
                        Dim s As String = RemoveStartEnd(m.Value, "[", "]")

                        'Split the parameter and value
                        Dim oParamVal As New ParameterValue(s, ":")

                        Select Case oParamVal.Parameter
                            Case "QS"
                                Dim sQS As String = GetQsValue(oParamVal.Value1)

                                If sQS = String.Empty Then
                                    'Return "" if querystring not found
                                    sOriginal = ReplaceToken(sOriginal, m.Value, String.Empty, Type)
                                Else
                                    sOriginal = ReplaceToken(sOriginal, m.Value, sQS, Type)
                                End If

                        End Select

                    End If

                End If

                'If the string still contains a [ try parsing path tokens
                If sOriginal.Contains("[") Then

                    sOriginal = PathTokenReplace(sOriginal)

                End If

            End If



            Return (sOriginal)

        End Function



        Private Function ReplaceSettingsToken(Original As String, SetType As SettingsType, TxtType As TextType) As String

            Dim strType As String = ""
            Dim TabId As Integer = -1

            Select Case SetType

                Case SettingsType.Application
                    strType = "Application"
                Case SettingsType.Portal
                    strType = "Portal"
                Case SettingsType.Tab
                    strType = "Page"
                    TabId = PortalSettings.ActiveTab.TabID

            End Select




            'Page settings, not a regex test for performance

            Dim strBase As String = String.Format("[{0}:Settings", strType)

            'Is there a setting in the token?
            If Original.Contains(strBase) Then


                ' Process Settings token

                'Create regex
                Dim strSettingRx As String = String.Format("\{0}\.(.*?)]", strBase)


                Dim rxSettings As Regex = New Regex(strSettingRx, RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant)

                'loop over matches

                For Each mtch As Match In rxSettings.Matches(Original)

                    Dim strName As String = mtch.Groups(1).Value
                    Dim strVal As String = GetSetting(SetType, TabId, strName)
                    Original = ReplaceToken(Original, mtch.Groups(0).Value, strVal, TxtType)

                Next



            End If

            Return Original

        End Function



        Private Function GetParentAlias() As String

            Return Request.Url.Host

        End Function

        Private Function GetAliasProtocol() As String

            Return Request.Url.Scheme

        End Function



        Private Function ReplaceToken(Original As String, Token As String, Value As String, Type As TextType) As String


            'Replace an indivifual token
            If Value Is Nothing Then
                Value = ""
            End If

            'Convert text to valid string for that type, when 
            Select Case Type

                Case TextType.CssClass

                    Value = CreateValidCssClass(Value)

                Case TextType.Text

                    'Do nothing

            End Select

            'Escape Token
            Token = Regex.Escape(Token)

            Return Regex.Replace(Original, Token, Value, RegexOptions.IgnoreCase)

        End Function

        Private Function MatchToken(Original As String, Token As String) As Boolean

            'Test if there is a match

            Return Regex.IsMatch(Original, Regex.Escape(Token), RegexOptions.IgnoreCase)

        End Function

#End Region


#Region "DNN Settings"

        Enum SettingsType

            Application
            Portal
            Tab

        End Enum




        Private Function GetSetting(Mode As SettingsType, TabId As Integer, PropertyName As String) As String



            Select Case Mode

                Case SettingsType.Application

                    'to be added later

                Case SettingsType.Portal

                    Return PortalController.GetPortalSetting(PropertyName, PortalSettings.PortalId, "")


                Case SettingsType.Tab

                    Dim oTabC As New TabController
                    Dim htTabSettings As Hashtable = oTabC.GetTabSettings(TabId)
                    Return htTabSettings(PropertyName).ToString



            End Select


            Return ""

        End Function


        Private Sub SaveSetting(Mode As SettingsType, PropertyName As String, Value As String)

            If ModuleControl Is Nothing Then

                Select Case Mode

                    Case SettingsType.Application

                        'to be added later

                    Case SettingsType.Portal

                        PortalController.UpdatePortalSetting(PortalSettings.PortalId, PropertyName, Value)


                    Case SettingsType.Tab

                        Dim oTabC As New TabController
                        oTabC.UpdateTabSetting(PortalSettings.ActiveTab.TabID, PropertyName, Value)


                End Select




            End If


        End Sub



#End Region

        Function GetTabData() As DotNetNuke.Entities.Tabs.TabInfo

            Dim oTab As DotNetNuke.Entities.Tabs.TabInfo = PortalSettings.ActiveTab

            If oTab.Title = "" Then oTab.Title = oTab.TabName
            If oTab.Description = "" Then oTab.Description = oTab.Title

            Return oTab

        End Function

        Function GetFullUrl() As String

            'Correct for Child portals
            Dim strProtocol As String = HttpContext.Current.Request.Url.Scheme
            Dim strAlias As String = HttpContext.Current.Request.Url.Host
            Dim strUrl As String = HttpContext.Current.Request.RawUrl

            Dim strUri As String = String.Format("{0}://{1}{2}", strProtocol, strAlias, strUrl)

            Return strUri


        End Function


#Region "MultiLanguage"



        Private Function GetLanguageLinks(TabId As Integer, Template As String) As String

            Dim strOut As String = String.Empty

            ' Check if Localization is enabled or not
            If LocalizationEnabled() Then

                Dim oLocalizedTabs As Dictionary(Of String, Locale) = DotNetNuke.Services.Localization.LocaleController.Instance.GetLocales(PortalSettings.PortalId)

                If Not oLocalizedTabs Is Nothing Then

                    For Each oLocale As KeyValuePair(Of String, Locale) In oLocalizedTabs

                        Dim oLoc As Locale = CType(oLocale.Value, Locale)

                        If oLoc.IsPublished Then

                            Dim strUrl = GetPageTranslated(TabId, oLoc.Code)

                            If strUrl <> String.Empty Then

                                Dim strLink As String = Template

                                ' Tokens
                                strLink = Regex.Replace(strLink, "\[URL\]", strUrl, RegexOptions.IgnoreCase)
                                strLink = Regex.Replace(strLink, "\[LOCALE:CODE\]", oLoc.Code, RegexOptions.IgnoreCase)


                                strOut &= strLink

                            End If

                        End If


                    Next

                End If

            End If

            Return strOut


        End Function

        ''' <summary>
        ''' Get the url of a page in the other language
        ''' </summary>
        ''' <param name="TabId"></param>
        ''' <param name="CultureCode"></param>
        ''' <returns></returns>
        Private Function GetPageTranslated(TabId As Integer, CultureCode As String) As String

            Dim oTabC As New TabController

            Dim oLocale As New DotNetNuke.Services.Localization.Locale
            oLocale = DotNetNuke.Services.Localization.LocaleController.Instance.GetLocale(CultureCode)


            If Not oLocale Is Nothing Then

                Dim oTab As DotNetNuke.Entities.Tabs.TabInfo = oTabC.GetTabByCulture(PortalSettings.ActiveTab.TabID, PortalSettings.PortalId, oLocale)

                If Not (oTab Is Nothing) AndAlso oTab.IsTranslated Then

                    Return oTab.FullUrl

                End If

            End If

            Return ""


        End Function


        Private Function LocalizationEnabled() As Boolean

            Return PortalController.GetPortalSettingAsBoolean("ContentLocalizationEnabled", PortalSettings.PortalId, False)

        End Function

#End Region

#Region "PageClasses"

        Private Sub ProcessPageClassTemplate(ByVal Template As String)

            'Process the template for the body CSS class
            Dim sOut As String = String.Empty

            Dim oTab As DotNetNuke.Entities.Tabs.TabInfo

            Template = SetCssClassCase(Template)

            'Process Single Tokens
            Template = ProcessTokens(Template, TextType.CssClass)


            sOut = String.Empty

            'Get Bc Name Class
            If Regex.IsMatch(Template, "\[BcName\]", RegexOptions.IgnoreCase) Then
                For Each oTab In PortalSettings.ActiveTab.BreadCrumbs
                    sOut &= " " & CreateValidCssClass("L" & oTab.Level & "_" & oTab.TabName)
                Next
                Template = Regex.Replace(Template, "\[BcName\]", sOut, RegexOptions.IgnoreCase)
            End If



            sOut = String.Empty

            'Get Bc Id Class
            If Regex.IsMatch(Template, "\[BcId\]", RegexOptions.IgnoreCase) Then
                For Each oTab In PortalSettings.ActiveTab.BreadCrumbs
                    sOut &= " " & CreateValidCssClass("Id" & oTab.TabID)
                Next
                Template = Regex.Replace(Template, "\[BcId\]", sOut, RegexOptions.IgnoreCase)
            End If

            sOut = String.Empty

            'Get Bc Order Class
            If Regex.IsMatch(Template, "\[BcNr\]", RegexOptions.IgnoreCase) Then
                For Each oTab In PortalSettings.ActiveTab.BreadCrumbs
                    sOut &= " " & CreateValidCssClass("L" & oTab.Level & "_" & "Nr" & GetTabOrder(oTab.TabID))
                Next
                Template = Regex.Replace(Template, "\[BcNr\]", sOut, RegexOptions.IgnoreCase)
            End If

            'Get Bc Level Class
            If Regex.IsMatch(Template, "\[BcLevel\]", RegexOptions.IgnoreCase) Then
                For Each oTab In PortalSettings.ActiveTab.BreadCrumbs
                    sOut &= " " & CreateValidCssClass("Level" & oTab.Level)
                Next
                Template = Regex.Replace(Template, "\[BcLevel\]", sOut, RegexOptions.IgnoreCase)
            End If


            'Get Bc Role Class
            If Regex.IsMatch(Template, "\[UserPageRoles\]", RegexOptions.IgnoreCase) Then
                Template = Regex.Replace(Template, "\[UserPageRoles\]", GetUserPageRolesClass, RegexOptions.IgnoreCase)
            End If



            WritePageClass(Template.Trim)


        End Sub


        Private Function GetUserPageRolesClass() As String

            'For all the roles that have view right on this page, get all roles the current user is a member of.
            'Please note that host is treaded as a member of all roles (so all roles will be reurned if you are logged in as Host)

            'Get Authorized roles for this page
            Dim sRoles As String = PortalSettings.ActiveTab.TabPermissions.ToString("VIEW")
            Dim sOut As String = String.Empty

            If sRoles > "" Then
                'Remove ; at the end
                sRoles = Regex.Replace(sRoles, ";$", "")

                'Split out all the roles
                For Each sRole As String In Regex.Split(sRoles, ";")
                    'Check if the user is a member of the Role
                    If UserController.GetCurrentUserInfo().IsInRole(sRole) Then
                        sOut &= " " & CreateValidCssClass("UPR_" & sRole)
                    End If
                Next
            End If

            Return sOut.Trim

        End Function


        Private Function GetPageType() As String
            'Get the type of the current page.

            Dim sTemplate As String = "PageType_{0}"

            Dim oTab As DotNetNuke.Entities.Tabs.TabInfo = PortalSettings.ActiveTab
            Dim iTabId As Integer = oTab.TabID

            Dim sOut As String = String.Empty

            'Home Tab?

            If iTabId = PortalSettings.HomeTabId Then
                sOut = ConcatCssClasses(sOut, "Home", sTemplate, False)
            End If

            If iTabId = PortalSettings.SplashTabId Then
                sOut = ConcatCssClasses(sOut, "Splash", sTemplate, False)
            End If


            ' Login page
            If iTabId = PortalSettings.LoginTabId Or Request.RawUrl.Contains("/ctl/Login/") Or Request.RawUrl.Contains("/Login?") Then
                sOut = ConcatCssClasses(sOut, "Login", sTemplate, False)
            ElseIf Request.RawUrl.Contains("/ctl/") Then

                'Edit page
                sOut = ConcatCssClasses(sOut, "Edit", sTemplate, False)
            End If


            If iTabId = PortalSettings.RegisterTabId Then
                sOut = ConcatCssClasses(sOut, "Register", sTemplate, False)
            End If

            If iTabId = PortalSettings.UserTabId Then
                sOut = ConcatCssClasses(sOut, "User", sTemplate, False)
            End If

            If iTabId = PortalSettings.SearchTabId Then
                sOut = ConcatCssClasses(sOut, "Search", sTemplate, False)
            End If

            'This is an admin or host tab
            If PortalSettings.ActiveTab.TabPermissions.ToString("VIEW") = "Administrators;" Or PortalSettings.ActiveTab.TabPermissions.ToString("VIEW") = "" Then
                sOut = ConcatCssClasses(sOut, "Admin", sTemplate, False)
            End If

            If sOut = "" Then
                sOut = ConcatCssClasses(sOut, "Normal", sTemplate, False)
            End If

            Return sOut

        End Function


        Private Function GetBrowserClass(max As Integer, BrowserFilter As String) As String
            'Get a class for the browser, max is the maximum version

            Dim sOut As New StringBuilder

            'Get the browser name
            Dim sUserBrowserName As String = GetBrowserName()


            If Regex.IsMatch(sUserBrowserName, BrowserFilter, RegexOptions.IgnoreCase) Then

                Dim iVersion As Integer = GetBrowserVersion().Major


                sOut.Append(String.Format("{0}{1}", sUserBrowserName, iVersion.ToString))


                If max > iVersion Then

                    Dim x As Integer

                    For x = iVersion + 1 To max

                        sOut.Append(String.Format(" lt-{0}{1}", sUserBrowserName, x))

                    Next

                End If

            End If




            Return sOut.ToString


        End Function




        Private Sub WritePageClass(ByVal sIn As String)

            Dim oBody As New HtmlGenericControl
            oBody = CType(Me.Page.FindControl("Body"), HtmlGenericControl)
            oBody.Attributes.Add("class", sIn.Trim)

        End Sub


        ''' <summary>
        ''' Add content after the body tag..
        ''' </summary>
        ''' <param name="sIn"></param>
        ''' <remarks></remarks>
        ''' 
        Private Sub ProcessAddToBody(sIn As String, Position As InjectPosition)

            Dim oBody As Control = Me.Page.FindControl("Body")

            If Not oBody Is Nothing Then

                Dim oLit As New Literal
                oLit.Text = ProcessTokens(sIn)

                If Position = InjectPosition.Top Then
                    oBody.Controls.AddAt(0, oLit)
                Else
                    oBody.Controls.AddAt(oBody.Controls.Count(), oLit)
                End If

            End If

        End Sub



        Private Function GetTabOrder(ByVal TabId As Integer) As Integer

            Dim iOrder As Integer = 0

            Dim oTabC As New DotNetNuke.Entities.Tabs.TabController
            Dim myTab As DotNetNuke.Entities.Tabs.TabInfo = oTabC.GetTab(TabId, PortalSettings.PortalId, True)

            Dim iParentId As Integer = 0
            If Not myTab Is Nothing Then
                iParentId = myTab.ParentId
            End If

            'DNN 5.2 minimum
            For Each oTab As DotNetNuke.Entities.Tabs.TabInfo In TabController.GetTabsByParent(iParentId, PortalSettings.PortalId)
                If Navigation.CanShowTab(oTab, False, False) And (oTab.CultureCode = String.Empty OrElse oTab.CultureCode = PortalSettings.CultureCode) Then
                    iOrder += 1
                    If oTab.TabID = TabId Then
                        Exit For
                    End If
                End If
            Next

            Return (iOrder)

        End Function


        Private Function SetCssClassCase(CssClass As String) As String

            Select Case CssClassCase.ToLower
                Case CssClassCases.lowercase.ToString.ToLower
                    CssClass = CssClass.ToLower
                Case CssClassCases.UPPERCASE.ToString.ToLower
                    CssClass = CssClass.ToUpper
                Case CssClassCases.PascalCase.ToString.ToLower
                    CssClass = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(CssClass.ToLower)
            End Select

            Return CssClass


        End Function



        Private Function CreateValidCssClass(ByVal sIn As String) As String


            sIn = SetCssClassCase(sIn)

            sIn = Regex.Replace(sIn, "^[^A-z]|[^A-z0-9]", "_")

            Return sIn

        End Function



        ''' <summary>
        ''' Concat Css Class strings
        ''' </summary>
        ''' <param name="Start">Base String to add to</param>
        ''' <param name="Add">String to add to Base</param>
        ''' <param name="Template">Tempalte for the String to add</param>
        ''' <param name="ValidateClass">Convert added String to valid CSS class</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ConcatCssClasses(Start As String, Add As String, Optional Template As String = "", Optional ValidateClass As Boolean = False) As String

            'Parse Template
            If Template > "" Then
                Add = String.Format(Template, Add)
            End If


            'Convert to valid CSS class if needed
            If ValidateClass Then
                Add = CreateValidCssClass(Add)
            End If

            If Start = String.Empty Then
                Return Add
            Else
                Return String.Format("{0} {1}", Start, Add)
            End If



        End Function



#End Region

#Region "Helper classes"

        Enum PageType

            Home
            Splash
            Login
            Register
            Account
            Search

        End Enum


        Public Class PortalInfo

            Private _oHome As PageInfo
            Public ReadOnly Property Home() As PageInfo
                Get
                    Return GetPageByType(PageType.Home)
                End Get
            End Property


            Private _oSplash As PageInfo
            Public ReadOnly Property Splash() As PageInfo
                Get
                    Return GetPageByType(PageType.Splash)
                End Get
            End Property


            Private _oLogin As PageInfo
            Public ReadOnly Property Login() As PageInfo
                Get
                    Return GetPageByType(PageType.Login)
                End Get
            End Property


            Private _oRegister As PageInfo
            Public ReadOnly Property Register() As PageInfo
                Get
                    Return GetPageByType(PageType.Register)
                End Get
            End Property


            Private _oAccount As PageInfo
            Public ReadOnly Property Account() As PageInfo
                Get
                    Return GetPageByType(PageType.Account)
                End Get
            End Property


            Private _oSearch As PageInfo
            Public ReadOnly Property Search() As PageInfo
                Get
                    Return GetPageByType(PageType.Search)
                End Get
            End Property


            Function GetPageByType(Type As PageType) As PageInfo

                Dim PortalSettings As DotNetNuke.Entities.Portals.PortalSettings = DotNetNuke.Common.Globals.GetPortalSettings()

                Dim Id As Integer = 0

                Select Case Type

                    Case PageType.Home

                        Id = PortalSettings.HomeTabId

                    Case PageType.Splash

                        Id = PortalSettings.SplashTabId

                    Case PageType.Login

                        Id = PortalSettings.LoginTabId

                    Case PageType.Register

                        Id = PortalSettings.RegisterTabId

                    Case PageType.Account

                        Id = PortalSettings.UserTabId

                    Case PageType.Search

                        Id = PortalSettings.SearchTabId

                End Select

                Dim oPage As New PageInfo(Id)

                Return (oPage)

            End Function

        End Class

        Public Class PageInfo

            ''' <summary>
            ''' 
            ''' </summary>
            ''' <remarks></remarks>
            Private _iId As Integer
            Public Property Id() As Integer
                Get
                    Return _iId
                End Get
                Set(ByVal value As Integer)
                    _iId = value
                End Set
            End Property


            Private _strUrl As String
            Public Property Url() As String
                Get
                    Return _strUrl
                End Get
                Set(ByVal value As String)
                    _strUrl = value
                End Set
            End Property


            Private _strName As String
            Public Property Name() As String
                Get
                    Return _strName
                End Get
                Set(ByVal value As String)
                    _strName = value
                End Set
            End Property


            Private _strStatus As String
            Public Property Status() As String
                Get
                    Return _strStatus
                End Get
                Set(ByVal value As String)
                    _strStatus = value
                End Set
            End Property




            Public Sub New(id As Integer)

                Me.Id = id

                Dim PortalSettings As DotNetNuke.Entities.Portals.PortalSettings = DotNetNuke.Common.Globals.GetPortalSettings()

                Dim oTabC As New TabController
                Dim oTab As DotNetNuke.Entities.Tabs.TabInfo = oTabC.GetTab(id, PortalSettings.PortalId, False)

                If Not oTab Is Nothing Then
                    Try
                        Me.Url = DotNetNuke.Common.Globals.NavigateURL(oTab.TabID)
                        Me.Name = oTab.TabName

                    Catch ex As Exception
                        Status = "Error-Getting-Url"

                    End Try

                Else
                    Status = "Error-Not-A-Valid-TabId"
                End If

                If Me.Id = -1 Then
                    Me.Url = "#"
                    Me.Name = "Page Not Set"
                End If

            End Sub


        End Class
#End Region


#Region "IsMobile"

        Function CheckMobile() As Boolean
            'Will return True if check is not needed.

            If _iIfMobile <> 0 Then
                Dim bCheckFor As Boolean = False
                'Get the correct return Value, make sure it will: 
                'Return True if IsMobile is set to True and this is a mobile browser.
                'Return True if IsMobile is set to False and this is NOT a mobile browser.
                If _iIfMobile = 1 Then bCheckFor = True


                'Detect if the browser is a mobile browser
                'The regex comes from as site that's not hijacked, but you can replace the regex if you want

                Dim u As String = Request.ServerVariables("HTTP_USER_AGENT")
                Dim b As New Regex(IfMobileRX1, RegexOptions.IgnoreCase)

                'In case the second regex was set to "" also match
                Dim bM2 As Boolean = False

                If IfMobileRX2 <> "" Then
                    Dim v As New Regex(IfMobileRX2, RegexOptions.IgnoreCase)
                    If v.IsMatch(Left(u, 4)) Then
                        bM2 = True
                    End If
                End If

                If b.IsMatch(u) Or bM2 Then
                    Return (bCheckFor)
                Else
                    Return (Not bCheckFor)
                End If
            Else
                Return True
            End If

        End Function


#End Region

    End Class

End Namespace



