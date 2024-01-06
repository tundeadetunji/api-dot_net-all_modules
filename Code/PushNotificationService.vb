Imports FirebaseAdmin
Imports FirebaseAdmin.Messaging
Imports Google.Apis.Auth.OAuth2
Imports iNovation.Code.General
Public Class PushNotificationService

    Private ReadOnly Property app

    'Private ReadOnly Property firebase_admin_sdk_json_file As String
    'Private ReadOnly Property credential

    Private ReadOnly Property firebase_token As String
    ''' <summary>
    ''' Required to send push notification to Android device.
    ''' 
    ''' Ṣe ètò àti ṣe ìfiránṣẹ́ sí ẹ̀rọ ìbánisọ̀rọ̀.
    ''' </summary>
    ''' <param name="firebase_admin_sdk_json_file"></param>
    ''' <param name="firebase_device_token"></param>
    Public Sub New(firebase_admin_sdk_json_file As String, firebase_device_token As String)
        Me.firebase_token = firebase_device_token
        app = FirebaseApp.Create(New AppOptions() With {.Credential = GoogleCredential.FromFile(firebase_admin_sdk_json_file)})
    End Sub

    Public Sub Push(title As String, message As String)

        'Dim messaging = FirebaseMessaging.GetMessaging(app)


        'Dim message = New MulticastMessage() With {
        '    .Tokens = New List(Of String) From {
        '        ReadText(firebase_token_file)
        '    },
        '    .Notification = New Notification() With {
        '        .Title = "Hello",
        '        .Body = "This is a test notification"
        '    }
        '}
        'FirebaseMessaging.GetMessaging(app).SendMulticastAsync(Message)
        FirebaseMessaging.GetMessaging(app).SendMulticastAsync(New MulticastMessage() With {
            .Tokens = New List(Of String) From {
                firebase_token
            },
            .Notification = New Notification() With {
                .Title = title,
                .Body = message
            }
        })

    End Sub

    ''Dim credential = GoogleCredential.FromFile(firebase_admin_sdk_json_file)



End Class
