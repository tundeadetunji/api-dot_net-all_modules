#Region "Imports"
Imports Web_Module.DW
Imports Web_Module.DataConnectionWeb
'Imports Statistics.Charts
#End Region
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web

''' <summary>
''' Traffic is private.
''' </summary>
Public Class Traffic

#Region "Fields"

#End Region

#Region "New"
	Public Sub New(con_string As String)
	End Sub
#End Region

#Region "Main"
	Public Shared ReadOnly Property default_inactive_since_date As DateTime = DateTime.Now.AddMonths(-3)
	Public Shared Function DeleteInactiveProfiles(Optional since As DateTime = Nothing, Optional category As Profile.ProfileAuthenticationOption = Profile.ProfileAuthenticationOption.Anonymous)
		Dim since__ As DateTime = since
		If since = Nothing Then since__ = default_inactive_since_date
		Return Profile.ProfileManager.DeleteInactiveProfiles(category, since__)
	End Function

	Public Shared Function NumberOfProfiles(category As Profile.ProfileAuthenticationOption)
		Return Profile.ProfileManager.GetNumberOfProfiles(category)
	End Function
	Public Shared Function GetProfileName()
		Dim p As New System.Web.Profile.ProfileBase()
		Return p.GetPropertyValue("UsernName")
	End Function
	Public Shared Function GetNumberOfVisits()
		Dim p As New System.Web.Profile.ProfileBase()
		Return p.GetPropertyValue("numberOfVisits")
	End Function
	Public Shared Function UserIsAnonymous() As Boolean
		Dim p As New System.Web.Profile.ProfileBase()
		Return p.GetPropertyValue("IsAnonymous")
	End Function
	Public Shared Function IsAnonymousUser() As Boolean
		Return UserIsAnonymous()
	End Function
#End Region

#Region "Helper Functions"

#End Region

End Class

''' <summary>
''' Cookies is Public.
''' </summary>
Public Class Cookies
#Region "Fields"
	Private page__ As Page

#End Region

#Region "New"

	Public Sub New(page_ As Page)
		page__ = page_
	End Sub
#End Region
#Region "Main"
	Public Function GetCookie(cookie_name As String, Optional key_ As String = Nothing) As String
		Dim return_ As String = ""
		If Not IsNothing(page__.Request.Cookies(cookie_name)) And key_ IsNot Nothing Then
			If page__.Request.Cookies(cookie_name).HasKeys Then
				return_ = page__.Request.Cookies(cookie_name)(key_)
			Else
				return_ = page__.Request.Cookies(cookie_name).Value
			End If
		End If
		Return return_
	End Function

	Public Function SetCookie(cookie_name As String, value_ As String, Optional key_ As String = Nothing, Optional is_persistent_cookie As Boolean = True, Optional domain_name As String = Nothing, Optional no_access_from_javascript As Boolean = False, Optional path_ As String = Nothing, Optional secure_ As Boolean = True) As Boolean
		Dim return_ As Boolean = False

		If Not IsNothing(page__.Response.Cookies(cookie_name)) And key_ IsNot Nothing Then
			If page__.Response.Cookies(cookie_name).HasKeys Then
				page__.Response.Cookies(cookie_name)(key_) = value_

				If is_persistent_cookie Then SetExpiry(cookie_name)
				If domain_name IsNot Nothing Then SetDomain(cookie_name, domain_name)
				SetHttpOnly(cookie_name, no_access_from_javascript)
				If path_ IsNot Nothing Then SetPath(cookie_name, path_)
				SetSecure(cookie_name, secure_)

				return_ = True
			Else
				page__.Response.Cookies(cookie_name).Value = value_
				return_ = True
			End If
		End If
		Return return_
	End Function

	Public Function DeleteCookie(cookie_name As String) As Boolean
		Dim return_ As Boolean = False
		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).Expires = DateTime.Now.AddDays(-1)
			return_ = True
		End If
		Return return_
	End Function

#End Region
#Region "Helper Functions"
	Public Function SetDomain(cookie_name As String, domain_name As String) As Boolean
		Dim return_ As Boolean = False
		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).Domain = domain_name
			return_ = True
		End If
		Return return_
	End Function
	Public Function GetDomain(cookie_name As String) As String
		Try
			Return page__.Response.Cookies(cookie_name).Domain
		Catch ex As Exception
		End Try
	End Function
	Public Function SetHttpOnly(cookie_name As String, Optional no_access_from_javascript As Boolean = True) As Boolean
		Dim return_ As Boolean = False
		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).HttpOnly = no_access_from_javascript
			return_ = True
		End If
		Return return_
	End Function
	Public Function GetHttpOnly(cookie_name As String) As Boolean
		Try
			Return page__.Response.Cookies(cookie_name).HttpOnly
		Catch ex As Exception
		End Try
	End Function
	Public Function SetPath(cookie_name As String, path_ As String) As Boolean
		Dim return_ As Boolean = False
		Dim prefx As String = "/", path__ As String '= path_
		If Mid(path_, 1, 1) <> "/" Then
			path__ = prefx & path_
		Else
			path__ = path_
		End If

		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).Path = path__
			return_ = True
		End If
		Return return_
	End Function
	Public Function GetPath(cookie_name As String) As String
		Try
			Return page__.Response.Cookies(cookie_name).Path
		Catch ex As Exception
		End Try
	End Function

	Public Function SetSecure(cookie_name As String, secure As Boolean) As Boolean
		Dim return_ As Boolean = False
		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).Secure = secure
			return_ = True
		End If
		Return return_
	End Function

	Public Function GetSecure(cookie_name As String) As Boolean
		Try
			Return page__.Response.Cookies(cookie_name).Secure
		Catch ex As Exception
		End Try
	End Function
	Public Function SetExpiry(cookie_name As String, Optional date_ As Date = Nothing) As Boolean
		Dim return_ As Boolean = False
		Dim date__ As Date = Now.AddYears(1)
		If date_ <> Nothing Then date__ = date_

		If Not IsNothing(page__.Response.Cookies(cookie_name)) Then
			page__.Response.Cookies(cookie_name).Expires = date__
			return_ = True
		End If
		Return return_
	End Function

	Public Function GetExpiry(cookie_name As String) As Date
		Try
			Return Date.Parse(page__.Response.Cookies(cookie_name).Expires)
		Catch ex As Exception
		End Try
	End Function

#End Region

End Class
