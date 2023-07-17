'from Assemblies
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text.RegularExpressions


'from Nuget
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Data.OleDb
Imports System.Collections.ObjectModel

Public Class General

#Region "Structures"

    Public Structure CustomApplicationInfo
        Public NickName As String
        Public DisplayName As String
        Public ProcessName As String
        Public Filename As String
        Public InstallLocation As String
    End Structure

#End Region

#Region "Enums"

    Public Enum SearchStringOperator
        OR_ = 0
        AND_ = 1
        NOT_ = 2
    End Enum

    Public Enum WhatToMonitor
        Video = 0
        Spreadsheet = 1
        PDF = 2
        Audio = 3
        Image = 4
        WordProcessorDocument = 5
        Code = 6 'w/ ides' e.g. .sln, .xml
        TextFile = 7
        Executable = 8
        Link = 9
        Others = 10
    End Enum
    Public Enum ShutDownOrNot
        Shutdown = 0
        Restart = 1
        Hibernate = 2
        Sleep = 3
        Lock = 4
    End Enum

    Public Enum HideFileOrFolderShould
        LeaveAsIs = 0
        SealTheResource = 1
        HideTheResource = 2
        MakeTheResourceVisible = 3
    End Enum

    Public Enum AppWindowState
        Hide = 0
        ShowNormal = 1
        ShowMinimized = 2
        ShowMaximized = 3
        ShowNormalNoActivate = 4
        Show = 5
        Minimize = 6
        ShowMinNoActivate = 7
        ShowNoActivate = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimized = 11
    End Enum
    Public Enum ShowElementAfterward
        Yes = 1
        No = 2
    End Enum

    Public Enum AlertIs
        alert
        danger
        success
        warning
    End Enum

    Public Enum Fix
        LineBreak
        LineBreakWithParagraph
    End Enum

    Public Enum WriteContentType
        None
        Fix
        Bootstrap
        Jumbotron
    End Enum

    Public Enum DeviceCapture
        SingleImage = 0
        MotionPicture = 1
        Video = 2
        ScreenSingleImage = 3
        ScreenMotionPicture = 4
        ScreenVideo = 5
        Audio = 6
    End Enum

    Public Enum SingularWord
        Message
        Mark
        Minute
        Day
        Student
        Month
        Year
        Hour
        Second
        Account
        Male
        Female
        User
        File
        Record
        Question
        Task
        Item
        Unit
        Room
        Client
        Match
        Invoice
        Sale
        Receipt
        Product
        Application
    End Enum
    Public Enum IDPattern
        Short_
        Short_DateOnly
        Short_DateTime
        Long_
        Long_DateTime

    End Enum
    Public Enum CommitDataTo
        CSV = 0
        Sequel = 1
        Access = 2
        Excel = 3
        Hibernate = 4
    End Enum

    Public Structure CommitDataTargetInfo
        Public commit_data_to As CommitDataTo
        Public filename As String
        'Public table As String
        'Public connection_string As String

    End Structure

    Public Enum OrderBy
        DESC
        ASC
    End Enum

    Public Enum TextCase
        Capitalize = 0
        UpperCase = 1
        LowerCase = 2
        None = 3
    End Enum
    Public Enum ValidAs
        FileName = 0
        FilePath_Regular = 1
        '		RegularString = 2
        FilePath_InJSON = 2
    End Enum
    Public Enum ListIsOf
        String_
        Object_
        Integer_
        Double_
        Short_
        Byte_

    End Enum
    Enum ControlsToCheck
        All
        Any
    End Enum
    Public Enum TimeValue
        ShortDate = 1
        LongDate = 2
        ShortTime = 3
        LongTime = 4
        Year = 5
        Month = 6
        Day = 7
        Hour = 8
        Minute = 9
        Second = 10
        Millisecond = 11
        DayOfWeek = 12
        DayOfYear = 13
    End Enum
    Public Enum PropertyToBind
        Text
        Image
        BackgroundImage
        Items
        Check
    End Enum

    Public Enum ReturnInfo
        AsArray = 0
        AsString = 1
        AsListOfString = 2
        AsListOfObject = 3
        AsCollectionOfString = 4
        AsCollectionOfObject = 5
        AsCustomApplicationInfo = 6
    End Enum
    Public Enum SideToReturn
        Left = 0
        Right = 1
        AsArray = 2
        AsListOfString = 3
        AsListToString = 4
        AsCustomApplicationInfo = 5
        AsCustomApplicationInfoDisplayName = 6
        AsCustomApplicationInfoProcessName = 7
        AsCustomApplicationInfoFilename = 8
        AsCustomApplicationInfoInstallLocation = 9
        Keys = 10
        Values = 11
    End Enum

    Public Shadows Enum FileType
        any
        gif
        jpg
        jpeg
        tif
        bmp
        png
        ico
        doc
        docx
        htm
        html
        xhtml
        mht
        aac
        aa3
        adt
        adts
        mp3
        wav
        mid
        midi
        wma
        _3gp
        _3gpp
        _3pg2
        m4v
        mp4
        mov
        mkv
        flv
        wmv
        avi
        mpg
        mpeg
        vob
        dat
        amv
        webm
        lnk
        txt
        pdf
        xls
        xlsx
        ppt
        pptx
        ppsx
        pps
        js
        java
        aspx
        asp
        py
        sql
        xml
        vb
        cs
        json
        sln
        bat
        md
        exe
        msi
        nuspec
        rar
    End Enum

    Enum FormatFor
        Web
        JavaScript
        Custom
        SQL
    End Enum
    Public Enum WrapWebContent
        AsIs = 0
        Div = 1
        P = 2
    End Enum
    Public Enum WriteAs
        AsIs = 0
        General = 1

    End Enum

    Public Enum Months
        January = 1
        February = 2
        March = 3
        April = 4
        May = 5
        June = 6
        July = 7
        August = 8
        September = 9
        October = 10
        November = 11
        December = 12
    End Enum

#End Region

#Region "Support"

    Public Shared ReadOnly Property StatusList(Optional IsUpdate As Boolean = False) As List(Of String)
        Get
            Dim l_ As New List(Of String)
            If IsUpdate = False Then l_.Add("Started")
            l_.Add("Done")
            l_.Add("Canceled")
            Return l_
        End Get
    End Property

    Public Shared ReadOnly Property GenderList() As List(Of String)
        Get
            Dim l_ As New List(Of String)
            l_.Add("Male")
            l_.Add("Female")
            Return l_
        End Get
    End Property

    Public Shared ReadOnly Property TitleList As List(Of String)
        Get
            Return {"Mr.", "Mrs.", "Ms."}.ToList
        End Get
    End Property

#Region "LGAsOfNigeriaList"
    Public Shared ReadOnly Property abia_ As Array = {"Aba North", "Aba South", "Arochukwu", "Bende", "Isiala Ngwa South", "Ikwuano", "Isiala", "Ngwa North", "Isukwuato", "Ukwa West", "Ukwa East", "Umuahia", "Umuahia South"}
    Public Shared ReadOnly Property adamawa_ As Array = {"Demsa", "Fufore", "Ganye", "Girei", "Gombi", "Jada", "Yola North", "Lamurde", "Madagali", "Maiha", "Mayo-Belwa", "Michika", "Mubi South", "Numna", "Shelleng", "Song", "Toungo", "Jimeta", "Yola South", "Hung"}
    Public Shared ReadOnly Property akwa_ibom As Array = {"Abak", "Eastern Obolo", "Eket", "Essien Upublic shared readonly property ", "Etimekpo", "Etinan", "Ibeno", "Ibesikpo Asutan", "Ibiono Ibom", "Ika", "Ikono", "Ikot Abasi", "Ikot Ekpene", "Ini", "Itu", "Mbo", "Mkpat Enin", "Nsit Ibom", "Nsit Ubium", "Obot Akara", "Okobo", "Onna", "Orukanam", "Oron", "Udung Uko", "Ukanafun", "Esit Eket", "Uruan", "Urue Offoung", "Oruko Ete", "Uyo"}
    Public Shared ReadOnly Property anambra_ As Array = {"Aguata,Anambra", "Anambra West", "Anaocha", "Awka South", "Awka North", "Ogbaru", "Onitsha South", "Onitsha North", "Orumba North", "Orumba South", "Oyi"}
    Public Shared ReadOnly Property bauchi_ As Array = {"Alkaleri", "Bauchi", "Bogoro", "Darazo", "Dass", "Gamawa", "Ganjuwa", "Giade", "Jama`are", "Katagum", "Kirfi", "Misau", "Ningi", "hira", "Tafawa Balewa", "Itas gadau", "Toro", "Warji", "Zaki", "Dambam"}
    Public Shared ReadOnly Property bayelsa_ As Array = {"Brass", "Ekeremor", "Kolok/Opokuma", "Nembe", "Ogbia", "Sagbama", "Southern Ijaw", "Yenagoa", "Membe"}
    Public Shared ReadOnly Property benue_ As Array = {"Ador", "Agatu", "Apa", "Buruku", "Gboko", "Guma", "Gwer East", "Gwer West", "Kastina-ala", "Konshisha", "Kwande", "Logo", "Makurdii", "Obi", "Ogbadibo", "Ohimini", "Oju", "Okpokwu", "Oturkpo", "Tarka", "Ukum", "Vandekya"}
    Public Shared ReadOnly Property borno_ As Array = {"Abadan", "Askira/Uba", "Bama", "Bayo", "Biu", "Chibok", "Damboa", "Dikwagubio", "Guzamala", "Gwoza", "Hawul", "Jere", "Kaga", "Kalka/Balge", "Konduga", "Kukawa", "Kwaya-ku", "Mafa", "Magumeri", "Maiduguri", "Marte", "Mobbar", "Monguno", "Ngala", "Nganzai", "Shani"}
    Public Shared ReadOnly Property cross_river As Array = {"Abia", "Akampa", "Akpabuyo", "Bakassi", "Bekwara", "Biase", "Boki", "Calabar South", "Etung", "Ikom", "Obanliku", "Obubra", "Obudu", "Odukpani", "Ogoja", "Ugep north", "Yala", "Yarkur"}
    Public Shared ReadOnly Property delta_ As Array = {"Aniocha South", "Anioha", "Bomadi", "Burutu", "Ethiope west", "Ethiope east", "Ika south", "Ika north east", "Isoko South", "Isoko north", "Ndokwa east", "Ndokwa west", "Okpe", "Oshimili north", "Oshimili south", "Patani", "Sapele", "Udu", "Ughelli south", "Ughelli north", "Ukwuani", "Uviwie", "Warri central", "Warri north", "Warri south"}
    Public Shared ReadOnly Property ebonyi_ As Array = {"Abakaliki", "Afikpo south", "Afikpo north", "Ebonyi", "Ezza", "Ezza south", "Ikwo", "Ishielu", "Ivo", "Ohaozara", "Ohaukwu", "Onicha", "Izzi"}
    Public Shared ReadOnly Property enugu_ As Array = {"Awgu", "Aninri", "Enugu east", "Enugu south", "Enugu north", "Ezeagu", "Igbo Eze north", "Igbi etiti", "Nsukka", "Oji river", "Undenu", "Uzo Uwani", "Udi"}
    Public Shared ReadOnly Property edo_ As Array = {"Akoko-Edo", "Egor", "Essann east", "Esan south east", "Esan central", "Esan west", "Etsako central", "Etsako east", "Etsako", "Orhionwon", "Ivia north", "Ovia south west", "Owan west", "Owan south", "Uhunwonde"}
    Public Shared ReadOnly Property ekiti_ As Array = {"Ado Ekiti", "Effon Alaiye", "Ekiti south west", "Ekiti west", "Ekiti east", "Emure/ise", "Orun", "Ido", "Osi", "Ijero", "Ikere", "Ikole", "Ilejemeje", "Irepodun", "Ise/Orun", "Moba", "Oye", "Aiyekire"}
    Public Shared ReadOnly Property federal_capital_territory As Array = {"Abaji", "Abuja Municipal", "Bwari", "Gwagwalada", "Kuje", "Kwali"}
    Public Shared ReadOnly Property gombe_ As Array = {"Akko", "Balanga", "Billiri", "Dukku", "Dunakaye", "Gombe", "Kaltungo", "Kwami", "Nafada/Bajoga", "Shomgom", "Yamaltu/Deba"}
    Public Shared ReadOnly Property imo_ As Array = {"Aboh-Mbaise", "Ahiazu-Mbaise", "Ehime-Mbaino", "Ezinhite", "Ideato North", "Ideato south", "Ihitte/Uboma", "Ikeduru", "Isiala", "Isu", "Mbaitoli", "Ngor Okpala", "Njaba", "Nwangele", "Nkwere", "Obowo", "Aguta", "Ohaji Egbema", "Okigwe", "Onuimo", "Orlu", "Orsu", "Oru west", "Oru", "Owerri", "Owerri North", "Owerri south"}
    Public Shared ReadOnly Property jigawa_ As Array = {"Auyo", "Babura", "Birnin-Kudu", "Birniwa", "Buji", "Dute", "Garki", "Gagarawa", "Gumel ", "Guri", "Gwaram", "Gwiwa", "Hadeji", "Jahun", "Kafin-Hausa", "Kaugama", "Kazaure", "Kirikisamma", "Birnin-magaji", "Maigatari", "Malamaduri", "Miga", "Ringim", "Roni", "Sule Tankarka", "Taura", "Yankwasi"}
    Public Shared ReadOnly Property kaduna_ As Array = {"Birnin Gwari", "Chukun", "Giwa", "Kajuru", "Igabi", "Ikara", "Jaba", "Jema`a", "Kachia", "Kaduna North", "Kaduna south", "Kagarok", "Kauru", "Kabau", "Kudan", "Kere", "Makarfi", "Sabongari", "Sanga", "Soba", "Zangon-Kataf", "Zaria"}
    Public Shared ReadOnly Property kano_ As Array = {"Ajigi", "Albasu", "Bagwai", "Bebeji", "Bichi", "Bunkure", "Dala", "Dambatta", "Dawakin kudu", "Dawakin tofa", "Doguwa", "Fagge", "Gabasawa", "Garko", "Garun mallam", "Gaya", "Gezawa", "Gwale", "Gwarzo", "Kano", "Karay", "Kibiya", "Kiru", "Kumbtso", "Kunch", "Kura", "Maidobi", "Makoda", "MInjibir Nassarawa", "Rano", "Rimin gado", "Rogo", "Shanono", "Sumaila", "Takai", "Tarauni", "Tofa", "Tsanyawa", "Tudunwada", "Ungogo", "Warawa", "Wudil"}
    Public Shared ReadOnly Property katsina_ As Array = {"Bakori", "Batagarawa", "Batsari", "Baure", "Bindawa", "Charanchi", "Dan-Musa", "Dandume", "Danja", "Daura", "Dutsi", "Dutsin `ma", "Faskar", "Funtua", "Ingawa", "Jibiya", "Kafur", "Kaita", "Kankara", "Kankiya", "Katsina", "Furfi", "Kusada", "Mai Adua", "Malumfashi", "Mani", "Mash", "Matazu", "Musawa", "Rimi", "Sabuwa", "Safana", "Sandamu", "Zango"}
    Public Shared ReadOnly Property kebbi_ As Array = {"Aliero", "Arewa Dandi", "Argungu", "Augie", "Bagudo", "Birnin Kebbi", "Bunza", "Dandi", "Danko", "Fakai", "Gwandu", "Jeda", "Kalgo", "Koko-besse", "Maiyaama", "Ngaski", "Sakaba", "Shanga", "Suru", "Wasugu", "Yauri", "Zuru"}
    Public Shared ReadOnly Property kogi_ As Array = {"Adavi", "Ajaokuta", "Ankpa", "Bassa", "Dekina", "Yagba east", "Ibaji", "Idah", "Igalamela", "Ijumu", "Kabba bunu", "Kogi", "Mopa muro", "Ofu", "Ogori magongo", "Okehi", "Okene", "Olamaboro", "Omala", "Yagba west"}
    Public Shared ReadOnly Property kwara_ As Array = {"Asa", "Baruten", "Ede", "Ekiti", "Ifelodun", "Ilorin south", "Ilorin west", "Ilorin east", "Irepodun", "Isin", "Kaiama", "Moro", "Offa", "Oke ero", "Oyun", "Pategi"}
    Public Shared ReadOnly Property lagos_ As Array = {"Agege", "Alimosho Ifelodun", "Alimosho", "Amuwo-Odofin", "Apapa", "Badagry", "Epe", "Eti-Osa", "Ibeju-Lekki", "Ifako/Ijaye", "Ikeja", "Ikorodu", "Kosofe", "Lagos Island", "Lagos Mainland", "Mushin", "Ojo", "Oshodi-Isolo", "Shomolu", "Surulere"}
    Public Shared ReadOnly Property nassarawa_ As Array = {"Akwanga", "Awe", "Doma", "Karu", "Keana", "Keffi", "Kokona", "Lafia", "Nassarawa", "Nassarawa/Eggon", "Obi", "Toto", "Wamba"}
    Public Shared ReadOnly Property niger_ As Array = {"Agaie", "Agwara", "Bida", "Borgu", "Bosso", "Chanchanga", "Edati", "Gbako", "Gurara", "Kitcha", "Kontagora", "Lapai", "Lavun", "Magama", "Mariga", "Mokwa", "Moshegu", "Muya", "Paiko", "Rafi", "Shiroro", "Suleija", "Tawa-Wushishi"}
    Public Shared ReadOnly Property ogun_ As Array = {"Abeokuta south", "Abeokuta north", "Ado-odo/otta", "Agbado south", "Agbado north", "Ewekoro", "Idarapo", "Ifo", "Ijebu east", "Ijebu north", "Ikenne", "Ilugun Alaro", "Imeko afon", "Ipokia", "Obafemi/owode", "Odeda", "Odogbolu", "Ogun Waterside", "Sagamu"}
    Public Shared ReadOnly Property ondo_ As Array = {"Akoko north", "Akoko north east", "Akoko south east", "Akoko south", "Akure north", "Akure", "Idanre", "Ifedore", "Ese odo", "Ilaje", "Ilaje oke-igbo", "Irele", "Odigbo", "Okitipupa", "Ondo", "Ondo east", "Ose", "Owo"}
    Public Shared ReadOnly Property osun_ As Array = {"Atakumosa west", "Atakumosa east", "Ayeda-ade", "Ayedire", "Bolawaduro", "Boripe", "Ede", "Ede north", "Egbedore", "Ejigbo", "Ife north", "Ife central", "Ife south", "Ife east", "Ifedayo", "Ifelodun", "Ilesha west", "Ilaorangun", "Ilesah east", "Irepodun", "Irewole", "Isokan", "Iwo", "Obokun", "Odo-otin", "Ola Oluwa", "Olorunda", "Oriade", "Orolu", "Osogbo"}
    Public Shared ReadOnly Property oyo_ As Array = {"Afijio", "Akinyele", "Atiba", "Atigbo", "Egbeda", "Ibadan North", "Ibadan North East", "Ibadan North West", "Ibadan south east", "Ibadan South West", "Ibarapa Central", "Ibarapa east", "Ibarapa north", "Ido", "Ifedapo", "Ifeloju", "Irepo", "Iseyin", "Itesiwaju", "Iwajowa", "Olorunshogo", "Kajola", "Lagelu", "Ogbomosho north", "Ogbomosho south", "Ogo Oluwa", "Oluyole", "Ona Ara", "Ore Lope", "Orire", "Oyo east", "Oyo west", "Saki east", "Saki west", "Surulere"}
    Public Shared ReadOnly Property plateau_ As Array = {"Barkin/ladi", "Bassa", "Bokkos", "Jos North", "Jos east", "Jos south", "Kanam", "Kiyom", "Langtang north", "Langtang south", "Mangu", "Mikang", "Pankshin", "Qua`an pan", "Shendam", "Wase"}
    Public Shared ReadOnly Property rivers_ As Array = {"Abua/Odial", "Ahoada west", "Akuku toru", "Andoni", "Asari toru", "Bonny", "Degema", "Eleme", "Emohua", "Etche", "Gokana", "Ikwerre", "Oyigbo", "Khana", "Obio/Akpor", "Ogba east/Edoni", "Ogu/bolo", "Okrika", "Omumma", "Opobo/Nkoro", "Portharcourt", "Tai"}
    Public Shared ReadOnly Property sokoto_ As Array = {"Binji", "Bodinga", "Dange/shuni", "Gada", "Goronyo", "Gudu", "Gwadabawa", "Illella", "Kebbe", "Kware", "Rabah", "Sabon-Birni", "Shagari", "Silame", "Sokoto south", "Sokoto north", "Tambuwal", "Tangaza", "Tureta", "Wamakko", "Wurno", "Yabo"}
    Public Shared ReadOnly Property taraba_ As Array = {"Akdo-kola", "Bali", "Donga", "Gashaka", "Gassol", "Ibi", "Jalingo", "K/Lamido", "Kurmi", "Lan", "Sardauna", "Tarum", "Ussa", "Wukari", "Yorro", "Zing"}
    Public Shared ReadOnly Property yobe_ As Array = {"Borsari", "Damaturu", "Fika", "Fune", "Geidam", "Gogaram", "Gujba", "Gulani", "Jakusko", "Karasuwa", "Machina", "Nagere", "Nguru", "Potiskum", "Tarmua", "Yunusari", "Yusufari", "Gashua"}
    Public Shared ReadOnly Property zamfara_ As Array = {"Anka", "Bukkuyum", "Dungudu", "Chafe", "Gummi", "Gusau", "Isa", "Kaura/Namoda", "Mradun", "Maru", "Shinkafi", "Talata/Mafara", "Zumi"}


#End Region

    Public Shared ReadOnly Property StatesOfNigeriaList As List(Of String)
        Get
            Dim l As New List(Of String)
            With l
                .Add("Abia")
                .Add("Adamawa")
                .Add("Akwa Ibom")
                .Add("Anambra")
                .Add("Bauchi")
                .Add("Bayelsa")
                .Add("Benue")
                .Add("Borno")
                .Add("Cross River")
                .Add("Delta")
                .Add("Ebonyi")
                .Add("Enugu")
                .Add("Edo")
                .Add("Ekiti")
                .Add("Federal Capital Territory")
                .Add("Gombe")
                .Add("Imo")
                .Add("Jigawa")
                .Add("Kaduna")
                .Add("Kano")
                .Add("Katsina")
                .Add("Kebbi")
                .Add("Kogi")
                .Add("Kwara")
                .Add("Lagos")
                .Add("Nassarawa")
                .Add("Niger")
                .Add("Ogun")
                .Add("Ondo")
                .Add("Osun")
                .Add("Oyo")
                .Add("Plateau")
                .Add("Rivers")
                .Add("Sokoto")
                .Add("Taraba")
                .Add("Yobe")
                .Add("Zamfara")
            End With
            Return l
        End Get
    End Property

    Public Shared ReadOnly Property CountriesList As List(Of String)
        Get
            Dim l As New List(Of String)
            With l
                .Add("Afghanistan")
                .Add("Albania")
                .Add("Algeria")
                .Add("Andorra")
                .Add("Angola")
                .Add("Antigua And Barbuda")
                .Add("Argentina")
                .Add("Armenia")
                .Add("Australia")
                .Add("Austria")
                .Add("Azerbaijan")
                .Add("Bahamas")
                .Add("Bahrain")
                .Add("Bangladesh")
                .Add("Barbados")
                .Add("Belarus")
                .Add("Belgium")
                .Add("Belize")
                .Add("Benin")
                .Add("Bhutan")
                .Add("Bolivia")
                .Add("Bosnia And Herzegovina")
                .Add("Botswana")
                .Add("Brazil")
                .Add("Brunei")
                .Add("Bulgaria")
                .Add("Burkina Faso")
                .Add("Burundi")
                .Add("Cambodia")
                .Add("Cameroon")
                .Add("Canada")
                .Add("Cape Verde")
                .Add("Central African Republic")
                .Add("Chad")
                .Add("Chile")
                .Add("China")
                .Add("Colombia")
                .Add("Comoros")
                .Add("Costa Rica")
                .Add("Cote d'Ivoire")
                .Add("Croatia")
                .Add("Cuba")
                .Add("Cyprus")
                .Add("Czech Republic")
                .Add("Democratic Republic Of Congo")
                .Add("Denmark")
                .Add("Djibouti")
                .Add("Dominica")
                .Add("Dominican Republic")
                .Add("East Timor")
                .Add("Ecuador")
                .Add("Egypt")
                .Add("El Salvador")
                .Add("Equitorial Guinea")
                .Add("Eritrea")
                .Add("Estonia")
                .Add("Ethiopia")
                .Add("Federal States Of Micronisia")
                .Add("Fiji")
                .Add("Finland")
                .Add("Former Yugoslav Republic Of Macedonia")
                .Add("France")
                .Add("Gabon")
                .Add("Georgia")
                .Add("Germany")
                .Add("Ghana")
                .Add("Greece")
                .Add("Grenada")
                .Add("Guatemala")
                .Add("Guinea")
                .Add("Guinea-Bissau")
                .Add("Guyana")
                .Add("Haiti")
                .Add("Honduras")
                .Add("Hungary")
                .Add("Iceland")
                .Add("India")
                .Add("Iran")
                .Add("Iraq")
                .Add("Ireland")
                .Add("Islamic Republic Of Mauritania")
                .Add("Israel")
                .Add("Italy")
                .Add("Jamaica")
                .Add("Japan")
                .Add("Jordan")
                .Add("Kazakhstan")
                .Add("Kenya")
                .Add("Kiribati")
                .Add("Kuwait")
                .Add("Kyrgyzstan")
                .Add("Laos")
                .Add("Latvia")
                .Add("Lebanon")
                .Add("Lesotho")
                .Add("Liberia")
                .Add("Libya")
                .Add("Liechtenstein")
                .Add("Lithuana")
                .Add("Luxembourg")
                .Add("Madagascar")
                .Add("Malawi")
                .Add("Malaysia")
                .Add("Maldives")
                .Add("Mali")
                .Add("Malta")
                .Add("Marshall Islands")
                .Add("Mauritius")
                .Add("Mexico")
                .Add("Moldova")
                .Add("Monaco")
                .Add("Mongolia")
                .Add("Montenegro")
                .Add("Morocco")
                .Add("Mozambique")
                .Add("Myanmar")
                .Add("Namibia")
                .Add("Nauru")
                .Add("Nepal")
                .Add("Netherlands")
                .Add("New Zealand")
                .Add("Nicaragua")
                .Add("Niger")
                .Add("Nigeria")
                .Add("North Korea")
                .Add("Norway")
                .Add("Oman")
                .Add("Pakistan")
                .Add("Panama")
                .Add("Papua New Guinea")
                .Add("Paraguay")
                .Add("Peru")
                .Add("Poland")
                .Add("Portugal")
                .Add("Qatar")
                .Add("Republic Of Congo")
                .Add("Republic Of Indonesia")
                .Add("Republic Of Singapore")
                .Add("Republic Of The Philippines")
                .Add("Republic Of Yemen")
                .Add("Romania")
                .Add("Russia")
                .Add("Rwanda")
                .Add("San Marino")
                .Add("Sao Tome And Principe")
                .Add("Saudi Arabia")
                .Add("Senegal")
                .Add("Serbia")
                .Add("Seychelles")
                .Add("Sierra Leone")
                .Add("Slovakia")
                .Add("Slovenia")
                .Add("Solomon Islands")
                .Add("Somalia")
                .Add("South Africa")
                .Add("South Korea")
                .Add("Spain")
                .Add("Sri Lanka")
                .Add("St Kitts-Nevis")
                .Add("St Lucia")
                .Add("St Vincent And The Grenadines")
                .Add("Sudan")
                .Add("Suriname")
                .Add("Swaziland")
                .Add("Sweden")
                .Add("Switzerland")
                .Add("Syria")
                .Add("Tajikistan")
                .Add("Tanzania")
                .Add("Thailand")
                .Add("The Gambia")
                .Add("Togo")
                .Add("Tonga")
                .Add("Trinidad And Tobago")
                .Add("Tunisia")
                .Add("Turkey")
                .Add("Turkmenistan")
                .Add("Tuvalu")
                .Add("Uganda")
                .Add("Ukraine")
                .Add("United Arab Emirates")
                .Add("United Kingdom")
                .Add("Uruguay")
                .Add("USA")
                .Add("Uzbekistan")
                .Add("Vanuatu")
                .Add("Vatican City")
                .Add("Venezuela")
                .Add("Vietnam")
                .Add("Western Samoa")
                .Add("Zaire")
                .Add("Zambia")
                .Add("Zimbabwe")
            End With
            Return l
        End Get
    End Property

#End Region

#Region "Sequel"
#Region "Enums"


    Public Enum LIKE_SELECT
        AND_
        OR_
    End Enum

    Public Enum Queries
        DeleteString_CONDITIONAL
        BuildSelectString_BETWEEN
        BuildSelectString_LIKE
        BuildUpdateString_CONDITIONAL
        BuildCountString_CONDITIONAL
        BuildCountString_GROUPED_CONDITIONAL
        BuildSelectString_CONDITIONAL
        BuildSelectString_DISTINCT
        BuildInsertString
        BuildUpdateString
        BuildSelectString
        BuildCountString
        BuildCountString_GROUPED
        BuildSumString_GROUPED
        BuildSumString_GROUPED_CONDITIONAL
        BuildAVGString_GROUPED
        BuildAVGString_GROUPED_CONDITIONAL
        BuildMinString_GROUPED
        BuildMinString_GROUPED_CONDITIONAL
        BuildTopString
        BuildMaxString_GROUPED
        BuildMaxString_GROUPED_CONDITIONAL
        BuildMaxString
        BuildAVGString_UNGROUPED
        BuildAVGString_UNGROUPED_CONDITIONAL
        BuildSumString_UNGROUPED
        BuildSumString_UNGROUPED_CONDITIONAL
        BuildCountString_UNGROUPED
        BuildCountString_UNGROUPED_CONDITIONAL
        BuildSelectString_DISTINCT_BETWEEN_CONDITIONAL
        BuildAVGString_GROUPED_BETWEEN
        BuildSumString_GROUPED_BETWEEN
        BuildCountString_GROUPED_BETWEEN
        BuildTopString_CONDITIONAL

        BuildSelectString_BETWEEN_Excel
        BuildSelectString_LIKE_Excel
        BuildSelectString_DISTINCT_Excel
        BuildInsertString_Excel
        BuildUpdateString_Excel
        BuildSelectString_Excel
        BuildCountString_Excel
        BuildCountString_GROUPED_Excel
        BuildSumString_GROUPED_Excel
        BuildAVGString_GROUPED_Excel
        BuildMinString_GROUPED_Excel
        BuildTopString_Excel
        BuildTopString_GROUPED_Excel
        BuildMaxString_GROUPED_Excel
        BuildMaxString_Excel
        BuildAVGString_UNGROUPED_Excel
        BuildSumString_UNGROUPED_Excel
        BuildCountString_UNGROUPED_Excel
        BuildAVGString_GROUPED_BETWEEN_Excel
        BuildSumString_GROUPED_BETWEEN_Excel
        BuildCountString_GROUPED_BETWEEN_Excel












    End Enum

#End Region

#Region "Query Strings - SQL Server, Access"



    ''' <summary>
    ''' Builds SQL Delete Query String.
    ''' </summary>
    ''' <param name="t_">Table to delete from.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to delete or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function DeleteString_CONDITIONAL(t_ As String, where_key_operator As Array) As String
        Dim where_keys As Array = where_key_operator
        Dim v As String = "DELETE "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            Else
                v = ""
            End If
            v &= ")"
        Else
            v = ""
        End If
        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Columns to select.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">Columns to consider with condition, appended automatically with _FROM or _TO (as must be manually appended to respective Where Keys when used with Display.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns></returns>
    Public Shared Function BuildSelectString_BETWEEN(t_ As String, Optional select_params As Array = Nothing, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
        End If
        If OrderByField IsNot Nothing Then
            v &= " ORDER BY " & OrderByField
        End If
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function
    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Columns to select.</param>
    ''' <param name="where_keys">Columns to apply condition on before selecting.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns></returns>
    Public Shared Function BuildSelectString_LIKE(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional like_operator As LIKE_SELECT = LIKE_SELECT.AND_, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & " LIKE '%' + @" & where_keys(j) & " + '%'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " " & like_operator.ToString.Replace("_", "") & " "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function


    ''' <summary>
    ''' Builds SQL Update Query String.
    ''' </summary>
    ''' <param name="t_">Table to update.</param>
    ''' <param name="update_keys">Fields to replace.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to update or not, followed by the operator to apply on the key.</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildUpdateString_CONDITIONAL(t_ As String, update_keys As Array, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "UPDATE " & t_ & " SET "

        For j As Integer = 0 To update_keys.Length - 1
            v &= update_keys(j) & "=@" & update_keys(j)
            If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("

                For k As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(k) & " " & operator_(where_keys(k + 1)) & " @" & where_keys(k)
                    If where_keys.Length > 1 And k <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
                v &= ")"
            End If
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count Query String. Useful for Line Chart (e.g. combined with Where clause)
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_CONDITIONAL(t_ As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT COUNT(*) "
        v &= " FROM " & t_

        Dim l As New List(Of String)
        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    If l.Contains(where_keys(j)) Then
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
                    Else
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    End If
                    l.Add(where_keys(j))
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function


    ''' <summary>
    ''' Builds SQL Count query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_GROUPED_CONDITIONAL(t_ As String, field_to_count As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT " & field_to_count & ", COUNT(*) "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_count
    End Function

    Private Shared Function operator_(operator__ As String) As String
        If operator__ = "" Then
            Return "="
        Else
            Return operator__
        End If
    End Function

    Private Shared Function BuildSelectString_COMPLEX(t_ As String, Optional select_params As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional where_keys_FOR_LIKE As Array = Nothing, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel)
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim where_keys_like As Array = where_keys_FOR_LIKE
        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Or where_keys_like IsNot Nothing Or where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            v &= " WHERE ("

            If where_keys IsNot Nothing Then
                If where_keys.Length > 0 Then
                    For j As Integer = 0 To where_keys.Length - 1 Step 2
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                        If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                            v &= " AND "
                        End If
                    Next
                End If
            End If

            If where_keys_like IsNot Nothing Then
                If where_keys_like.Length > 0 Then
                    If where_keys Is Nothing Then
                        v &= ""
                    Else
                        v &= " AND "
                    End If
                    For j As Integer = 0 To where_keys_like.Length - 1
                        v &= where_keys_like(j) & " LIKE '%' + @" & where_keys_like(j) & " + '%'"
                        If where_keys_like.Length > 1 And j <> where_keys_like.Length - 1 Then
                            v &= " AND "
                        End If
                    Next
                End If
            End If

            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
                If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                    If where_keys Is Nothing And where_keys_like Is Nothing Then
                        v &= ""
                    ElseIf where_keys IsNot Nothing Or where_keys_like IsNot Nothing Then
                        v &= " AND "
                    End If
                    For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                        v &= "" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO"
                        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                            v &= " AND "
                        End If
                    Next
                End If
            End If

            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to select or not, followed by the operator to apply on the key.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildSelectString_CONDITIONAL(t_ As String, Optional select_params As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        Dim l As New List(Of String)
        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    If l.Contains(where_keys(j)) Then
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
                    Else
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    End If
                    l.Add(where_keys(j))
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String with DISTINCT. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <returns>String.</returns>

    Public Shared Function BuildSelectString_DISTINCT(t_ As String, select_params As Array, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT DISTINCT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Insert Query String.
    ''' </summary>
    ''' <param name="t_">Table to insert into.</param>
    ''' <param name="insert_keys">Columns to insert into.</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildInsertString(t_ As String, insert_keys As Array, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "INSERT INTO " & t_ & " ("

        For i As Integer = 0 To insert_keys.Length - 1
            v &= insert_keys(i)
            If insert_keys.Length > 1 And i <> insert_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        v &= ") VALUES ("

        For j As Integer = 0 To insert_keys.Length - 1
            v &= "@" & insert_keys(j)
            If insert_keys.Length > 1 And j <> insert_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        v &= ")"
        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Update Query String.
    ''' </summary>
    ''' <param name="t_">Table to update.</param>
    ''' <param name="update_keys">Fields to replace.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to update or not.</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildUpdateString(t_ As String, update_keys As Array, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "UPDATE " & t_ & " SET "

        For j As Integer = 0 To update_keys.Length - 1
            v &= update_keys(j) & "=@" & update_keys(j)
            If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("

                For k As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(k) & "=@" & where_keys(k)
                    If where_keys.Length > 1 And k <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
                v &= ")"
            End If
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildSelectString(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count Query String.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString(t_ As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String

        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT COUNT(*) "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_GROUPED(t_ As String, field_to_count As String, field_to_group As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_count & ", COUNT(*) "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL Sum query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_GROUPED(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL Sum query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sum or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_sum As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_GROUPED(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sort for AVG or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL MIN query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find min.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_min_on">Field to sort for MIN.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MIN or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMinString_GROUPED(t_ As String, field_to_group As String, field_to_apply_min_on As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", MIN(" & field_to_apply_min_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL MIN query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find min.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_min_on">Field to sort for MIN.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sort for MIN or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMinString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_min_on As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT " & field_to_group & ", MIN(" & field_to_apply_min_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL Select Top Query String.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildTopString(t_ As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
        If select_keys IsNot Nothing Then
            v = "SELECT TOP " & Val(rows_to_select) & " "
            With select_keys
                For i As Integer = 0 To .Length - 1
                    v &= select_keys(i)
                    If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
                Next
            End With
            v &= " "
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v
    End Function

    Public Shared Function BuildTopString_GROUPED(t_ As String, field_to_group As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
        If select_keys IsNot Nothing Then
            v = "SELECT TOP " & Val(rows_to_select) & " "
            With select_keys
                For i As Integer = 0 To .Length - 1
                    v &= select_keys(i)
                    If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
                Next
            End With
            v &= " "
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v & " GROUP BY " & field_to_group
    End Function

    Public Shared Function BuildTopString_CONDITIONAL(t_ As String, Optional select_keys As Array = Nothing, Optional where_key_operator As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim select_params As Array = select_keys
        Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
        If select_keys IsNot Nothing Then
            v = "SELECT TOP " & Val(rows_to_select) & " "
            With select_keys
                For i As Integer = 0 To .Length - 1
                    v &= select_keys(i)
                    If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
                Next
            End With
            v &= " "
        End If
        v &= " FROM " & t_

        Dim l As New List(Of String)
        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    If l.Contains(where_keys(j)) Then
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j) & "_" & l.LastIndexOf(where_keys(j))
                    Else
                        v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    End If
                    l.Add(where_keys(j))
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL MAX query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find MAX.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_max_on">Field to sort for MAX.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MAX or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMaxString_GROUPED(t_ As String, field_to_group As String, field_to_apply_max_on As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", MAX(" & field_to_apply_max_on & ") "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL MAX query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find MAX.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_max_on">Field to sort for MAX.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sort for MAX or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMaxString_GROUPED_CONDITIONAL(t_ As String, field_to_group As String, field_to_apply_max_on As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT " & field_to_group & ", MAX(" & field_to_apply_max_on & ") "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    Public Shared Function BuildMaxString(t_ As String, Max_Field As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT MAX (" & Max_Field & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_UNGROUPED(t_ As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT AVG(" & field_to_apply_avg_on & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function


    ''' <summary>
    ''' Builds SQL AVG query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sort for AVG or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_UNGROUPED_CONDITIONAL(t_ As String, field_to_apply_avg_on As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT AVG(" & field_to_apply_avg_on & ")"
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function
    ''' <summary>
    ''' Builds SQL Sum query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_UNGROUPED(t_ As String, field_to_sum As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT SUM(" & field_to_sum & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Sum query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to sum or not, followed by the operator to apply on the key.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_UNGROUPED_CONDITIONAL(t_ As String, field_to_sum As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT SUM(" & field_to_sum & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function
    ''' <summary>
    ''' Builds SQL Count query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_UNGROUPED(t_ As String, field_to_count As String, Optional where_keys As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT COUNT(" & field_to_count & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "=@" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function


    ''' <summary>
    ''' Builds SQL Count query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_key_operator">Fields to check condition on, to decide to count or not, followed by the operator to apply on the key.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_UNGROUPED_CONDITIONAL(t_ As String, field_to_count As String, Optional where_key_operator As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim where_keys As Array = where_key_operator
        Dim v As String = "SELECT COUNT(" & field_to_count & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1 Step 2
                    v &= where_keys(j) & " " & operator_(where_keys(j + 1)) & " @" & where_keys(j)
                    If where_keys.Length > 1 And j <> where_keys.Length - 2 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String with DISTINCT. Useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">Columns to consider with condition, appended with _FROM or _TO.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildSelectString_DISTINCT_BETWEEN_CONDITIONAL(t_ As String, select_params As Array, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT DISTINCT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If

        End If
        If OrderByField IsNot Nothing Then
            v &= " ORDER BY " & OrderByField
        End If
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function
    ''' <summary>
    ''' Builds SQL AVG query string, grouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
        End If
        Return v & " GROUP BY " & field_to_group
    End Function
    ''' <summary>
    ''' Builds SQL SUM query string, grouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If

        End If
        Return v & " GROUP BY " & field_to_group
    End Function
    ''' <summary>
    ''' Builds SQL Count query string, grouped - useful for Line chart. When declaring keys, field must not contain _From or _To. When declaring keys_values, field must have _From and _To, e.g. RecordDate_From, RecordDate_To
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_count">Field to count.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">single field name, e.g. RecordDate</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_GROUPED_BETWEEN(t_ As String, field_to_group As String, field_to_count As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Sequel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", COUNT(" & field_to_count & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_FROM AND @" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & "_TO)"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If

        End If
        Return v & " GROUP BY " & field_to_group
    End Function


#End Region

#Region "Query Strings - Excel"

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Columns to select.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">Columns to consider with condition, appended automatically with _FROM or _TO (as must be manually appended to respective Where Keys when used with Display.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns></returns>
    Public Shared Function BuildSelectString_BETWEEN_Excel(t_ As String, Optional select_params As Array = Nothing, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional where_values As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN '" & where_values(j) & "' AND '" & where_values(j) & "')"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
        End If
        If OrderByField IsNot Nothing Then
            v &= " ORDER BY " & OrderByField
        End If
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Columns to select.</param>
    ''' <param name="where_keys">Columns to apply condition on before selecting.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns></returns>
    Public Shared Function BuildSelectString_LIKE_Excel(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional like_operator As LIKE_SELECT = LIKE_SELECT.AND_, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & " LIKE '%' + " & where_values(j) & " + '%'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " " & like_operator.ToString.Replace("_", "") & " "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v

    End Function

    ''' <summary>
    ''' Builds SQL Select Query String with DISTINCT. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildSelectString_DISTINCT_Excel(t_ As String, select_params As Array, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT DISTINCT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds Excel Insert Query String. USE WITH CAUTION, BECAUSE OF SQL INJECTION VULNERABILITY
    ''' </summary>
    ''' <param name="t_">Table to insert into.</param>
    ''' <param name="insert_keys">Columns to insert into.</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildInsertString_Excel(t_ As String, insert_keys As Array, insert_values As Array, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "INSERT INTO " & t_ & " ("

        For i As Integer = 0 To insert_keys.Length - 1
            v &= insert_keys(i)
            If insert_keys.Length > 1 And i <> insert_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        v &= ") VALUES ("

        For j As Integer = 0 To insert_keys.Length - 1
            v &= "'" & insert_values(j) & "'"
            If insert_keys.Length > 1 And j <> insert_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        v &= ")"
        Return v
    End Function

    ''' <summary>
    ''' ToDo/WIP
    ''' </summary>
    ''' <param name="t_">Table to update.</param>
    ''' <param name="update_keys">Fields to replace.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to update or not.</param>
    ''' <returns>String.</returns>
    Private Shared Function BuildUpdateString_Excel(t_ As String, update_keys As Array, Optional where_keys As Array = Nothing, Optional update_values As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "UPDATE " & t_ & " SET "

        For j As Integer = 0 To update_keys.Length - 1
            v &= update_keys(j) & "='" & update_values(j) & "'"
            If update_keys.Length > 1 And j <> update_keys.Length - 1 Then
                v &= ", "
            End If
        Next

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("

                For k As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(k) & "='" & where_values(k) & "'"
                    If where_keys.Length > 1 And k <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
                v &= ")"
            End If
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Select Query String. Suitable for Reader. To use count instead, use BuildCountString.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="select_params">Fields to select.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildSelectString_Excel(t_ As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT "

        If select_params IsNot Nothing Then
            For i As Integer = 0 To select_params.Length - 1
                v &= select_params(i)
                If select_params.Length > 1 And i <> select_params.Length - 1 Then
                    v &= ", "
                End If
            Next
        Else
            v &= " *"
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString
        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count Query String.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_Excel(t_ As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String

        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT COUNT(*) "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_GROUPED_Excel(t_ As String, field_to_count As String, field_to_group As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_count & ", COUNT(*) "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL Sum query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_GROUPED_Excel(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_GROUPED_Excel(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL MIN query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find min.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_min_on">Field to sort for MIN.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MIN or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMinString_GROUPED_Excel(t_ As String, field_to_group As String, field_to_apply_min_on As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", MIN(" & field_to_apply_min_on & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL Select Top Query String.
    ''' </summary>
    ''' <param name="t_">Table to select from.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to select or not.</param>
    ''' <param name="OrderByField">Column to order by.</param>
    ''' <param name="order_by">Ascending or Descending</param>
    ''' <returns>String.</returns>
    Public Shared Function BuildTopString_Excel(t_ As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional rows_to_select As Long = 1, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
        If select_keys IsNot Nothing Then
            v = "SELECT TOP " & Val(rows_to_select) & " "
            With select_keys
                For i As Integer = 0 To .Length - 1
                    v &= select_keys(i)
                    If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
                Next
            End With
            v &= " "
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v
    End Function

    Public Shared Function BuildTopString_GROUPED_Excel(t_ As String, field_to_group As String, Optional select_keys As Array = Nothing, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional rows_to_select As Long = 10, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.DESC, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT TOP " & Val(rows_to_select) & " * "
        If select_keys IsNot Nothing Then
            v = "SELECT TOP " & Val(rows_to_select) & " "
            With select_keys
                For i As Integer = 0 To .Length - 1
                    v &= select_keys(i)
                    If select_keys.Length > 1 And i <> select_keys.Length - 1 Then v &= ", "
                Next
            End With
            v &= " "
        End If
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If
        If OrderByField IsNot Nothing Then v &= " ORDER BY " & OrderByField
        If OrderByField IsNot Nothing Then v &= " " & order_by.ToString

        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL MAX query string, grouped - useful for Pie chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find MAX.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_max_on">Field to sort for MAX.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for MAX or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildMaxString_GROUPED_Excel(t_ As String, field_to_group As String, field_to_apply_max_on As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT " & field_to_group & ", MAX(" & field_to_apply_max_on & ") "
        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v & " GROUP BY " & field_to_group
    End Function

    Public Shared Function BuildMaxString_Excel(t_ As String, Max_Field As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT MAX (" & Max_Field & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sort for AVG or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_UNGROUPED_Excel(t_ As String, field_to_apply_avg_on As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT AVG(" & field_to_apply_avg_on & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Sum query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to sum or not.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_UNGROUPED_Excel(t_ As String, field_to_sum As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT SUM(" & field_to_sum & ") "

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL Count query string, ungrouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="where_keys">Fields to check equality condition on, to decide to count or not.</param>
    ''' <param name="field_to_count">Field to count and group.</param>
    ''' <returns>String</returns>
    Public Shared Function BuildCountString_UNGROUPED_Excel(t_ As String, field_to_count As String, Optional where_keys As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If
        Dim v As String = "SELECT COUNT(" & field_to_count & ")"

        v &= " FROM " & t_

        If where_keys IsNot Nothing Then
            If where_keys.Length > 0 Then
                v &= " WHERE ("
                For j As Integer = 0 To where_keys.Length - 1
                    v &= where_keys(j) & "='" & where_values(j) & "'"
                    If where_keys.Length > 1 And j <> where_keys.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
            v &= ")"
        End If

        Return v
    End Function

    ''' <summary>
    ''' Builds SQL AVG query string, grouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_apply_avg_on">Field to sort for AVG.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
    ''' <returns>String</returns>
    Public Shared Function BuildAVGString_GROUPED_BETWEEN_Excel(t_ As String, field_to_group As String, field_to_apply_avg_on As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", AVG(" & field_to_apply_avg_on & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN '" & where_values(j) & "' AND '" & where_values(j) & "')"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If
        End If
        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' Builds SQL SUM query string, grouped - useful for Line chart.
    ''' </summary>
    ''' <param name="t_">Table for fields to find AVG.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_sum">Field to sum.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO"></param>
    ''' <returns>String</returns>
    Public Shared Function BuildSumString_GROUPED_BETWEEN_Excel(t_ As String, field_to_group As String, field_to_sum As String, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional where_values As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", SUM(" & field_to_sum & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN '" & where_values(j) & "' AND '" & where_values(j) & "')"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If

        End If
        Return v & " GROUP BY " & field_to_group
    End Function

    ''' <summary>
    ''' ToDo/WIP
    ''' Builds SQL Count query string, grouped - useful for Line chart. When declaring keys, field must not contain _From or _To. When declaring keys_values, field must have _From and _To, e.g. RecordDate_From, RecordDate_To
    ''' </summary>
    ''' <param name="t_">Table for fields to count.</param>
    ''' <param name="field_to_group">Field to group.</param>
    ''' <param name="field_to_count">Field to count.</param>
    ''' <param name="where_keys_UNDERSCORE_FROM_UNDERSCORE_TO">single field name, e.g. RecordDate</param>
    ''' <returns>String</returns>
    Private Shared Function BuildCountString_GROUPED_BETWEEN_Excel(t_ As String, field_to_group As String, field_to_count As String, from As Object, to_ As Object, Optional where_keys_UNDERSCORE_FROM_UNDERSCORE_TO As Array = Nothing, Optional db As CommitDataTo = CommitDataTo.Excel) As String
        If db = CommitDataTo.Excel Then
            Dim t_temp = "[" & t_ & "$]"
            t_ = t_temp
        End If

        Dim v As String = "SELECT " & field_to_group & ", COUNT(" & field_to_count & ") "
        v &= " FROM " & t_

        If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO IsNot Nothing Then
            If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 0 Then
                v &= " WHERE "
                For j As Integer = 0 To where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1
                    v &= "(" & where_keys_UNDERSCORE_FROM_UNDERSCORE_TO(j) & " BETWEEN '" & from & "' AND '" & to_ & "')"
                    If where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length > 1 And j <> where_keys_UNDERSCORE_FROM_UNDERSCORE_TO.Length - 1 Then
                        v &= " AND "
                    End If
                Next
            End If

        End If
        Return v & " GROUP BY " & field_to_group
    End Function


#End Region

#Region "Retrieval"
    Public Enum Optimization
        FaultTolerant = 0
        AsIs = 1
        ByteArray = 2
    End Enum

    Public Shared Function QImage(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing)
        Return QData(query_, connection_string, parameters_keys_values_, Optimization.ByteArray)
    End Function


    ''' <summary>
    ''' Executes SQL statement for a single value.
    ''' </summary>
    ''' <param name="query_">SQL Query. You can use BuildSelectString, BuildUpdateString, BuildInsertString, BuildCountString, BuildTopString instead.</param>
    ''' <param name="connection_string">Connection String.</param>
    ''' <param name="parameters_keys_values_">Parameters.</param>
    ''' <param name="return_type_is">Return string by default. To return image, choose ByteArray</param>
    ''' <returns></returns>
    Public Shared Function QData(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional return_type_is As Optimization = Optimization.FaultTolerant)
        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = parameters_keys_values_

        Dim con As SqlConnection = New SqlConnection(connection_string)
        Dim cmd As SqlCommand = New SqlCommand(query_, con)

        If select_parameter_keys_values IsNot Nothing Then
            For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
            Next
        End If

        Try
            Using con
                con.Open()
                If return_type_is = Optimization.AsIs Then
                    '					Return CType(cmd.ExecuteScalar(), String).ToString
                    Return cmd.ExecuteScalar()
                ElseIf return_type_is = Optimization.FaultTolerant Then
                    Return CType(cmd.ExecuteScalar(), String).ToString
                ElseIf return_type_is = Optimization.ByteArray Then
                    Dim file() As Byte = CType(cmd.ExecuteScalar(), Byte())
                    Return file
                End If
            End Using
        Catch
        End Try

    End Function

    ''' <summary>
    ''' ToDo/WIP
    ''' </summary>
    ''' <param name="query_"></param>
    ''' <param name="connection_string"></param>
    ''' <returns></returns>
    Private Shared Function QDataFromExcel(query_ As String, connection_string As String)
        Dim table As DataTable = iNovation.Code.Sequel.QDataTableFromExcel(query_, connection_string)
        Dim result = table.Rows(0).Item(0).ToString
        Return result
    End Function


    ''' <summary>
    ''' Query should retrieve only one field, then this returns the whole of the columns'values in a list.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <returns></returns>
    Public Shared Function QList(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As List(Of Object)
        Dim l As New List(Of Object)

        Dim select_parameter_keys_values() = {}
        select_parameter_keys_values = select_parameter_keys_values_

        Try

            Dim connection As New SqlConnection(connection_string)
            Dim sql As String = query

            Dim Command = New SqlCommand(sql, connection)
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If

            Dim da As New SqlDataAdapter(Command)
            Dim dt As New DataTable
            da.Fill(dt)

            With dt
                For i = 0 To .Rows.Count - 1
                    l.Add(.Rows(i).Item(0).ToString)
                Next
            End With
        Catch
        End Try

        Return l
    End Function

    'Public Shared Function QDataWORKING(query_ As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional return_type_is As Want = Want.Value)
    '	Dim select_parameter_keys_values() = {}
    '	select_parameter_keys_values = parameters_keys_values_

    '	Dim con As SqlConnection = New SqlConnection(connection_string)
    '	Dim cmd As SqlCommand = New SqlCommand(query_, con)

    '	If select_parameter_keys_values IsNot Nothing Then
    '		For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '			cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '		Next
    '	End If

    '	Try
    '		Using con
    '			con.Open()
    '			If return_type_is = Want.Value Then
    '				Return CType(cmd.ExecuteScalar(), String).ToString
    '			ElseIf return_type_is = Want.Display Then
    '			End If
    '		End Using
    '	Catch
    '	End Try

    'End Function

    ''' <summary>
    ''' Checks if a field exists w/ or w/o specified condition.
    ''' </summary>
    ''' <param name="t_">Table to perform operation on.</param>
    ''' <param name="connection_string">Connection String.</param>
    ''' <param name="where_keys">List of parameters.</param>
    ''' <param name="where_keys_values">List of parameters and their values.</param>
    ''' <returns>True if it exists, False otherwise.</returns>

    Public Shared Function QExists(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing) As Boolean
        Return QCount(t_, connection_string, where_keys, where_keys_values) > 0
    End Function

    Public Shared Function QExists_CONDITIONAL(t_ As String, connection_string As String, Optional where_key_operator As Array = Nothing, Optional where_keys_values As Array = Nothing) As Boolean
        Return QData(BuildCountString_CONDITIONAL(t_, where_key_operator), connection_string, where_keys_values) > 0
    End Function

    ''' <summary>
    ''' Highest value of Max_Field for given choices w/ or w/o specified condition.
    ''' </summary>
    ''' <param name="t_">Table to perform operation on.</param>
    ''' <param name="connection_string">Connection String.</param>
    ''' <param name="where_keys">List of parameters.</param>
    ''' <param name="where_keys_values">List of parameters and their values.</param>
    ''' <param name="Max_Field">Field to use as maximum.</param>
    ''' <returns>Value of Max_Field.</returns>
    Public Shared Function QMax(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing, Optional Max_Field As String = Nothing)
        Return QData(BuildMaxString(t_, Max_Field, where_keys), connection_string, where_keys_values)
    End Function

    Public Shared Function QCount(t_ As String, connection_string As String, Optional where_keys As Array = Nothing, Optional where_keys_values As Array = Nothing)
        Return QData(BuildCountString(t_, where_keys), connection_string, where_keys_values)
    End Function
    Public Shared Function QCount_CONDITIONAL(t_ As String, connection_string As String, Optional where_key_operator As Array = Nothing, Optional where_keys_values As Array = Nothing)
        Return QData(BuildCountString_CONDITIONAL(t_, where_key_operator), connection_string, where_keys_values_CONDITIONAL(where_keys_values))

    End Function
#End Region
#Region "Retrieval - Raw"
    'Public Shared Function QTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As DataTable

    '    Try
    '        Using con As New SqlConnection(connection_string)
    '            Using cmd As New SqlCommand(query)
    '                Using sda As New SqlDataAdapter()
    '                    cmd.Connection = con
    '                    If select_parameter_keys_values IsNot Nothing Then
    '                        For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '                            cmd.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '                        Next
    '                    End If
    '                    sda.SelectCommand = cmd
    '                    Using dt As New DataTable
    '                        sda.Fill(dt)
    '                        Return dt
    '                    End Using
    '                End Using
    '            End Using
    '        End Using
    '    Catch ex As Exception
    '    End Try

    'End Function

#End Region

#Region "Code"
    Public Structure QParameters
        Public operation As Queries
        Public table As String
        Public SelectColumns As Array
        Public WhereKeys As Array
        Public InsertKeys As Array
        Public WhereOperators As Array
        Public MaxField As String
        Public OrderByField As String
        Public OrderRecordsBy As OrderBy
        Public LikeSelect As LIKE_SELECT
        Public TopRowsToSelect As Long
        Public UpdateKeys As Array
    End Structure
    Public Enum QOutput
        QData
        Display
        QString
        Commit
        QExists
        QCount
        QCount_Conditional
        BindProperty
    End Enum
    Public Enum Output
        Web
        Desktop
    End Enum

    Public Shared Function QString(q_parameters As QParameters, Optional q_output As QOutput = QOutput.QData, Optional connection_string_code_variable As String = Nothing, Optional output_ As Output = Output.Web) As String
        Dim operation As Queries = q_parameters.operation, table As String = q_parameters.table, SelectColumns As Array = q_parameters.SelectColumns, WhereKeys As Array = q_parameters.WhereKeys, InsertKeys As Array = q_parameters.InsertKeys, WhereOperators As Array = q_parameters.WhereOperators, MaxField As String = q_parameters.MaxField, OrderByField As String = q_parameters.OrderByField, OrderRecordsBy As OrderBy = q_parameters.OrderRecordsBy, LikeSelect As LIKE_SELECT = q_parameters.LikeSelect, TopRowsToSelect As Long = q_parameters.TopRowsToSelect, UpdateKeys As Array = q_parameters.UpdateKeys
        Dim r As String

        If operation = Queries.BuildUpdateString Then
            'BuildUpdateString(table__, {uk}, {wk})
            r = "BuildUpdateString(""" & table & """" 'WhereKeys)"
            If UpdateKeys IsNot Nothing Then
                With UpdateKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & UpdateKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            r &= ")"
        End If

        If operation = Queries.BuildTopString Then
            'BuildTopString(table, {sk}, {wk}, rows_to_select:=1, "order_by_field", OrderBy.ASC)
            r = "BuildTopString(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            r &= ", " & TopRowsToSelect
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildTopString_CONDITIONAL Then
            'BuildTopString_CONDITIONAL(table, {sk}, {wko}, rows_to_select:=1, "order_by_field", OrderBy.ASC)
            r = "BuildTopString_CONDITIONAL(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If TopRowsToSelect > 0 Then
                r &= ", " & TopRowsToSelect
            Else
                r &= ", 1"
            End If
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildSelectString_LIKE Then
            'BuildSelectString_LIKE(table__, {sk}, {wk}, "order_by_field", OrderBy.ASC, LIKE_SELECT.AND_)
            r = "BuildSelectString_LIKE(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            If LikeSelect <> Nothing Then
                r &= ", LIKE_SELECT." & LikeSelect.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildSelectString_DISTINCT Then
            'BuildSelectString_DISTINCT(table__, {SelectColumns}, {wk})
            r = "BuildSelectString_DISTINCT(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            r &= ")"
        End If

        If operation = Queries.BuildSelectString_CONDITIONAL Then
            'BuildSelectString_CONDITIONAL(table__, {SelectColumns}, {wko}, "order_by_field", OrderBy.ASC)
            r = "BuildSelectString_CONDITIONAL(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildSelectString_BETWEEN Then
            'BuildSelectString_BETWEEN(table__, {SelectColumns}, {WhereKeys}, "order_by_field", OrderBy.ASC)
            'SELECT MyAdminMedia FROM MyAdminMedia WHERE (RecordSerial BETWEEN @RecordSerial_FROM AND @RecordSerial_TO)
            r = "BuildSelectString_BETWEEN(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildSelectString Then
            'BuildSelectString(table__, {SelectColumns}, {WhereKeys}, "order_by_field", OrderBy.ASC)
            r = "BuildSelectString(""" & table & """" 'WhereKeys)"
            If SelectColumns IsNot Nothing Then
                With SelectColumns
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & SelectColumns(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            Else
                r &= ", Nothing"
            End If
            If OrderByField IsNot Nothing Then
                r &= ", """ & OrderByField & """"
                r &= ", OrderBy." & OrderRecordsBy.ToString
            End If
            r &= ")"
        End If

        If operation = Queries.BuildMaxString Then
            'BuildMaxString(table, where_keys, Max_Field)
            r = "BuildMaxString(""" & table & """" 'WhereKeys)"
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}, "
                End With
            Else
                r &= ", Nothing, "
            End If
            r &= """" & MaxField & """)"
        End If

        If operation = Queries.BuildInsertString Then
            'BuildInsertString(table, InsertKeys)
            r = "BuildInsertString(""" & table & """" 'WhereKeys)"
            If InsertKeys IsNot Nothing Then
                With InsertKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & InsertKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            End If
            r &= ")"
        End If

        If operation = Queries.BuildCountString_CONDITIONAL Then
            'BuildCountString_CONDITIONAL(table, WhereKeyOperator)
            r = "BuildCountString_CONDITIONAL(""" & table & """" 'WhereKeys)"
            If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            End If
            r &= ")"
        End If

        If operation = Queries.BuildCountString Then
            'BuildCountString(table, WhereKeys)
            r = "BuildCountString(""" & table & """" 'WhereKeys)"
            If WhereKeys IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            End If
            r &= ")"
        End If

        If operation = Queries.DeleteString_CONDITIONAL Then
            'DeleteString_CONDITIONAL(table, {wko})
            r = "DeleteString_CONDITIONAL(""" & table & """" 'WhereKeys)"

            If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
                With WhereKeys
                    r &= ", {"
                    For i As Integer = 0 To .Length - 1
                        r &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
                        If i < .Length - 1 Then r &= ", "
                    Next
                    r &= "}"
                End With
            End If
            r &= ")"
        End If

        If q_output = QOutput.QString Then Return r

        Dim wko As String = ""
        If WhereKeys IsNot Nothing And WhereOperators IsNot Nothing Then
            With WhereKeys
                wko = "{"
                For i As Integer = 0 To .Length - 1
                    wko &= """" & WhereKeys(i) & """, """ & WhereOperators(i) & """"
                    If i < .Length - 1 Then
                        wko &= ", "
                    End If
                Next
                wko &= "}"
            End With
        End If

        Dim wk As String = ""
        If WhereKeys IsNot Nothing Then
            With WhereKeys
                wk = "{"
                For i As Integer = 0 To .Length - 1
                    wk &= """" & WhereKeys(i) & """"
                    If i < .Length - 1 Then
                        wk &= ", "
                    End If
                Next
                wk &= "}"
            End With
        End If

        Dim wkv_b As String = ""
        If WhereKeys IsNot Nothing Then
            With WhereKeys
                wkv_b = "{"
                For i As Integer = 0 To .Length - 1
                    wkv_b &= """" & WhereKeys(i) & "_FROM"", """", """ & WhereKeys(i) & "_TO"", """""
                    If i < .Length - 1 Then
                        wkv_b &= ", "
                    End If
                Next
                wkv_b &= "}"
            End With
        End If

        Dim wkv As String = ""
        Dim l_wkv_conditional As New List(Of String)
        If WhereKeys IsNot Nothing Then
            If operation = Queries.BuildSelectString_CONDITIONAL Or operation = Queries.BuildCountString_CONDITIONAL Or operation = Queries.BuildTopString_CONDITIONAL Then
                With WhereKeys
                    wkv = "{"
                    For i As Integer = 0 To .Length - 1
                        If l_wkv_conditional.Contains(WhereKeys(i)) Then
                            wkv &= """" & WhereKeys(i) & "_" & l_wkv_conditional.LastIndexOf(WhereKeys(i)) & """, """" "
                        Else
                            wkv &= """" & WhereKeys(i) & """, """" "
                        End If
                        l_wkv_conditional.Add(WhereKeys(i))
                        If i < .Length - 1 Then
                            wkv &= ", "
                        End If
                    Next
                    wkv &= "}"
                End With
            Else
                With WhereKeys
                    wkv = "{"
                    For i As Integer = 0 To .Length - 1
                        wkv &= """" & WhereKeys(i) & """, """" "
                        If i < .Length - 1 Then
                            wkv &= ", "
                        End If
                    Next
                    wkv &= "}"
                End With
            End If
        End If

        Dim ikv As String = ""
        If InsertKeys IsNot Nothing Then
            With InsertKeys
                ikv = "{"
                For i As Integer = 0 To .Length - 1
                    ikv &= """" & InsertKeys(i) & """, """" "
                    If i < .Length - 1 Then
                        ikv &= ", "
                    End If
                Next
                ikv &= "}"
            End With
        End If

        Dim ukv As String = ""
        If UpdateKeys IsNot Nothing Then
            With UpdateKeys
                ukv = "{"
                For i As Integer = 0 To .Length - 1
                    ukv &= """" & UpdateKeys(i) & """, """" "
                    If i < .Length - 1 Then
                        ukv &= ", "
                    End If
                Next
                If WhereKeys IsNot Nothing Then
                    With WhereKeys
                        ukv &= ", "
                        For i As Integer = 0 To .Length - 1
                            ukv &= """" & WhereKeys(i) & """, """" "
                            If i < .Length - 1 Then
                                ukv &= ", "
                            End If
                        Next
                    End With
                End If
                ukv &= "}"
            End With
        End If

        Dim q As String = ""

        If q_output = QOutput.BindProperty Then
            If output_ = Output.Desktop Then
                'BindProperty(New Object, PropertyToBind.Items, r, "con", {kv}, "dataTextField", "dataValueField")
                q = "BindProperty(control_here, PropertyToBind.Items, " & r & ", connection_string_here"
                If connection_string_code_variable IsNot Nothing Then q = "BindProperty(control_here, PropertyToBind.Items, " & r & ", " & connection_string_code_variable
                If wkv.Length > 0 Then
                    q &= ", " & wkv
                Else
                    q &= ", Nothing"
                End If
                q &= ", """ & SelectColumns(0) & """"
                q &= ")"
            Else
                'BindProperty(control_here, r, "con", {kv}, "dtf")
                q = "BindProperty(control_here, " & r & ", connection_string_here"
                If connection_string_code_variable IsNot Nothing Then q = "BindProperty(control_here, " & r & ", " & connection_string_code_variable
                If wkv.Length > 0 Then
                    q &= ", " & wkv
                Else
                    q &= ", Nothing"
                End If
                q &= ", """ & SelectColumns(0) & """"
                q &= ")"
            End If
            Return q
        End If

        If q_output = QOutput.QCount Then
            'QCount(table, "conn", {where_keys()}, {wherekv})
            q = "QCount(""" & table & """, connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "QCount(""" & table & """, " & connection_string_code_variable

            If wk.Length > 0 Then
                q &= ", " & wk
            End If
            If wkv.Length > 0 Then
                q &= ", " & wkv
            End If
            q &= ")"
            Return q
        End If

        If q_output = QOutput.QCount_Conditional Then
            'QCount_CONDITIONAL(table, "con", {wko}, {wkv})
            q = "QCount_CONDITIONAL(""" & table & """, connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "QCount_CONDITIONAL(""" & table & """, " & connection_string_code_variable

            If wko.Length > 0 Then
                q &= ", " & wko
            End If
            If wkv.Length > 0 Then
                q &= ", " & wkv
            End If
            q &= ")"
            Return q
        End If

        If q_output = QOutput.QExists Then
            'QExists(table, server_con, {wk}, {wkv})
            q = "QExists(""" & table & """, connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "QExists(""" & table & """, " & connection_string_code_variable

            If wk.Length > 0 Then
                q &= ", " & wk
            End If
            If wkv.Length > 0 Then
                q &= ", " & wkv
            End If
            q &= ")"
            Return q
        End If

        If q_output = QOutput.QData Then
            'QData(r, server_con, {KV})
            q = "QData(" & r & ", connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "QData(" & r & ", " & connection_string_code_variable
            If WhereKeys IsNot Nothing Then
                If operation <> Queries.BuildSelectString_BETWEEN Then
                    q &= ", " & wkv
                Else
                    q &= ", " & wkv_b
                End If
            End If
            q &= ")"
            Return q
        End If

        If q_output = QOutput.Display Then
            'Display(grid, q)
            q = "Display(grid_here, " & r & ", connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "Display(grid_here, " & r & ", " & connection_string_code_variable
            If WhereKeys IsNot Nothing Then
                If operation <> Queries.BuildSelectString_BETWEEN Then
                    q &= ", " & wkv
                Else
                    q &= ", " & wkv_b
                End If
            End If
            q &= ")"
            Return q
        End If

        If q_output = QOutput.Commit Then
            q = "CommitSequel(" & r & ", connection_string_here"
            If connection_string_code_variable IsNot Nothing Then q = "CommitSequel(" & r & ", " & connection_string_code_variable
            'If WhereKeys IsNot Nothing Then
            If operation = Queries.BuildInsertString Then
                q &= ", " & ikv
            ElseIf operation = Queries.BuildUpdateString Then
                q &= ", " & ukv
            ElseIf operation = Queries.BuildUpdateString_CONDITIONAL Then
                q &= ", " & ukv
            End If
            q &= ")"
            Return q
        End If


    End Function

#End Region

#Region "Support Functions"
    Private Shared Function where_keys_values_CONDITIONAL(where_keys_values As Array) As Array
        Dim k As Array = where_keys(where_keys_values)
        Dim v As Array = where_values(where_keys_values)
        Dim l As New List(Of String)
        Dim lkv As New List(Of String)
        With k
            For i As Integer = 0 To .Length - 1
                If l.Contains(k(i)) Then
                    lkv.Add(k(i) & "_" & l.LastIndexOf(k(i)))
                    lkv.Add(v(i))
                Else
                    lkv.Add(k(i))
                    lkv.Add(v(i))
                End If
                l.Add(k(i))
            Next
        End With

        Return lkv.ToArray
    End Function
    Private Shared Function where_keys_emptyValues_CONDITIONAL(where_keys As Array) As Array
        Dim k As Array = where_keys(where_keys)
        Dim l As New List(Of String)
        Dim lkv As New List(Of String)
        With k
            For i As Integer = 0 To .Length - 1
                If l.Contains(k(i)) Then
                    lkv.Add(k(i) & "_" & l.LastIndexOf(k(i)))
                    lkv.Add("""""")
                Else
                    lkv.Add(k(i))
                    lkv.Add("""""")
                End If
                l.Add(k(i))
            Next
        End With

        Return lkv.ToArray
    End Function

    Public Shared Function where_keys(where_keys_values As Array) As Array
        If where_keys_values Is Nothing Then Return Nothing
        If where_keys_values.Length < 1 Then Return Nothing
        Dim l As New List(Of String)
        With where_keys_values
            For i As Integer = 0 To .Length - 2 Step 2
                l.Add(where_keys_values(i))
            Next
        End With
        Return l.ToArray
    End Function

    Public Shared Function where_values(where_keys_values As Array) As Array
        If where_keys_values Is Nothing Then Return Nothing
        If where_keys_values.Length < 1 Then Return Nothing
        Dim l As New List(Of String)
        With where_keys_values
            For i As Integer = 0 To .Length - 2 Step 2
                l.Add(where_keys_values(i + 1))
            Next
        End With
        Return l.ToArray
    End Function
#End Region

#Region "CU"


    ''' <summary>
    ''' Commits record to SQL Server database by default, or to MS Access database if DB_Is_SQL_ is set to false. Same as CommitRecord.
    ''' </summary>
    ''' <param name="query">The SQL query.</param>
    ''' <param name="connection_string">The server connection string.</param>
    ''' <param name="parameters_keys_values_">Values to put in table.</param>
    ''' <returns>True if successful, False if not.</returns>
    'Public Shared Function CommitSequel(query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing, Optional DB_Is_SQL_ As Boolean = True) As Boolean
    '    Dim select_parameter_keys_values() = {}
    '    select_parameter_keys_values = parameters_keys_values_

    '    If DB_Is_SQL_ = True Then
    '        CommitSQLRecord(query, connection_string, select_parameter_keys_values)
    '        Return True
    '        Exit Function
    '    End If

    '    Try
    '        Dim insert_query As String = query
    '        Using insert_conn As New OleDbConnection(connection_string)
    '            Using insert_comm As New OleDbCommand()
    '                With insert_comm
    '                    .Connection = insert_conn
    '                    .CommandText = insert_query
    '                    If select_parameter_keys_values IsNot Nothing Then
    '                        For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '                            .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '                        Next
    '                    End If
    '                End With
    '                Try
    '                    insert_conn.Open()
    '                    insert_comm.ExecuteNonQuery()
    '                Catch ex As Exception
    '                End Try
    '            End Using
    '        End Using
    '        Return True
    '    Catch ex As Exception
    '    End Try

    'End Function
    'Private Shared Function CommitSQLRecord(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As Boolean
    '    Dim select_parameter_keys_values() = {}
    '    select_parameter_keys_values = select_parameter_keys_values_
    '    Try
    '        Dim insert_query As String = query
    '        Using insert_conn As New SqlConnection(connection_string)
    '            Using insert_comm As New SqlCommand()
    '                With insert_comm
    '                    .Connection = insert_conn
    '                    .CommandTimeout = 0
    '                    .CommandType = CommandType.Text
    '                    .CommandText = insert_query
    '                    If select_parameter_keys_values IsNot Nothing Then
    '                        For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
    '                            .Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
    '                        Next
    '                    End If
    '                End With
    '                Try
    '                    insert_conn.Open()
    '                    insert_comm.ExecuteNonQuery()
    '                Catch ex As Exception
    '                End Try
    '            End Using
    '        End Using
    '        Return True
    '    Catch ex As Exception
    '    End Try

    '    '		Dim Entries_Insert As String = "INSERT INTO ENTRIES (EntryBy, ID, Category, [Description], Flag, [Title], Entry, DateAdded, TimeAdded, TitleID, Picture, PictureExtension, Topic) VALUES (@EntryBy, @ID, @Category, [@Description], @Flag, [@Title], @Entry, @DateAdded, @TimeAdded, @TitleID, @Picture, @PictureExtension, @Topic)"
    '    '		Dim entries_parameters_() = {"EntryBy", TitleBy.Text.Trim, "ID", EntryID.Text.Trim, "Category", Category.Text.Trim, "[Description]", Description.Text.Trim, "Flag", cFlag.Text.Trim, "[Title]", EntryTitle.Text.Trim, "Entry", NewEntry.Text.Trim, "DateAdded", date_, "TimeAdded", time_, "TitleID", TitleID.Text.Trim, "Picture", stream.GetBuffer(), "PictureExtension", PictureExtension.Text.Trim, "Topic", Topic.Text.Trim}
    '    '		d.CommitRecord(Entries_Insert, a_con, entries_parameters_)

    'End Function

#End Region

#End Region


#Region "Sequel - Internal"

    Public Shared Function GetDataTable(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing) As DataTable

        Try
            Dim connection As New SqlConnection(connection_string)
            Dim sql As String = query

            Dim Command = New SqlCommand(sql, connection)
            If select_parameter_keys_values IsNot Nothing Then
                For i As Integer = 0 To select_parameter_keys_values.Length - 1 Step 2
                    Command.Parameters.AddWithValue(select_parameter_keys_values(i), select_parameter_keys_values(i + 1))
                Next
            End If

            Dim da As New SqlDataAdapter(Command)
            Dim dt As New DataTable

            'da = New SqlDataAdapter(Command)
            'dt = New DataTable
            da.Fill(dt)
            Return dt
        Catch
            Throw New Exception("Table is empty")
        End Try

    End Function

#End Region

#Region "Sequel - API"
#Region "CSV"
    ''' <summary>
    ''' Creates CSV file from table View
    ''' </summary>
    ''' <param name="where_to_place_data"></param>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="parameters_keys_values_"></param>
    Public Shared Sub CommitData(where_to_place_data As CommitDataTargetInfo, query As String, connection_string As String, Optional parameters_keys_values_ As Array = Nothing)
        If Mid(query, 1, Len("select")).ToLower <> "select" And Mid(query, 1, Len("update")).ToLower <> "update" Then
            Return
        End If

        Dim data As DataTable = GetDataTable(query, connection_string, parameters_keys_values_)
        Dim header As String = ""
        Dim content As String = ""

        With data
            'header
            For h = 0 To .Columns.Count - 1
                header &= .Columns(h).ColumnName
                If h <> .Columns.Count Then header &= ","
            Next

            For i = 0 To .Rows.Count - 1
                For j = 0 To .Columns.Count - 1
                    'content
                    content &= .Rows(i).Item(j)
                    If j <> .Columns.Count Then content &= ","
                Next
                content &= vbCrLf
            Next
        End With

        WriteText(where_to_place_data.filename, header.Trim & vbCrLf & content.Trim)
    End Sub

#End Region

#End Region


#Region "JSON - API"

    ''' <summary>
    ''' Converts a row to JSON Object. query must return one row only.
    ''' </summary>
    ''' <param name="query"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <returns>JObject</returns>
    Public Shared Function RowToJSON(query As String, connection_string As String, Optional select_parameter_keys_values_ As Array = Nothing) As JObject
        Try
            Dim dt As DataTable = GetDataTable(query, connection_string, select_parameter_keys_values_)
            Dim d As New Dictionary(Of String, Object)
            With dt
                For i = 0 To .Rows.Count - 1
                    For j = 0 To .Columns.Count - 1
                        d.Add(.Columns.Item(j).ColumnName, .Rows(i).Item(j))
                    Next
                Next
            End With
            Dim jsonObject As JObject = JObject.Parse(JsonConvert.SerializeObject(d))
            Return jsonObject

        Catch ex As Exception

        End Try
    End Function
    ''' <summary>
    ''' Converts all rows (i.e. table) to objects and returns a list (JArray) containing them
    ''' </summary>
    ''' <param name="t_"></param>
    ''' <param name="connection_string"></param>
    ''' <param name="pk"></param>
    ''' <param name="select_params"></param>
    ''' <param name="where_keys"></param>
    ''' <param name="select_parameter_keys_values_"></param>
    ''' <param name="OrderByField"></param>
    ''' <param name="order_by"></param>
    ''' <returns>JArray</returns>
    Public Shared Function RowsToJSON(t_ As String, connection_string As String, pk As String, Optional select_params As Array = Nothing, Optional where_keys As Array = Nothing, Optional select_parameter_keys_values_ As Array = Nothing, Optional OrderByField As String = Nothing, Optional order_by As OrderBy = OrderBy.ASC) As JArray
        Dim overall_q = BuildSelectString(t_, select_params, where_keys, OrderByField, order_by)
        Dim dt As DataTable = GetDataTable(overall_q, connection_string, select_parameter_keys_values_)
        Dim l As New List(Of JObject)
        For i = 0 To dt.Rows.Count - 1
            Dim pk_query = BuildSelectString(t_, {pk})
            Dim pk_value As Object = dt.Rows(i).Item(pk)
            Dim pk_kv = {pk, pk_value}
            Dim q = BuildSelectString(t_, Nothing, {pk})
            ''Dim dt_ As DataTable = GetDataTable(q, connection_string, {pk, pk_value})
            Dim jobject As JObject = RowToJSON(q, connection_string, pk_kv)
            l.Add(jobject)
        Next
        Dim j As JArray = JArray.Parse(JsonConvert.SerializeObject(l))
        Return j
    End Function
    Public Enum OutputAs
        JObject = 0
        String_ = 1
    End Enum
    Public Shared Function SampleJSON(query As String, connection_string As String, Optional select_parameter_keys_values As Array = Nothing, Optional table_ As String = "whatever", Optional output_as As OutputAs = OutputAs.String_) As Object
        Dim t As DataTable = GetDataTable(query, connection_string, select_parameter_keys_values)
        Dim s As String = "{""table"": """ & table_ & """, ""shoulddisplay"": ""all"", ""shouldupdate"": ""id"", ""shoulddelete"": ""enabled"", "
        Dim r
        With t
            For i = 0 To .Columns.Count - 1
                s &= """" & .Columns.Item(i).ColumnName & """: "
                If .Columns.Item(i).DataType = GetType(Boolean) Then
                    s &= True.ToString.ToLower
                ElseIf .Columns.Item(i).DataType = GetType(Date) Then
                    s &= """" & Date.Now & """"
                ElseIf .Columns.Item(i).DataType = GetType(Byte) Then
                    s &= """byte_value"""
                ElseIf .Columns.Item(i).DataType = GetType(TimeSpan) Then
                    s &= """timespan_value"""
                Else
                    s &= """" & .Columns.Item(i).ColumnName.ToLower & """"
                End If
                If i <> .Columns.Count - 1 Then
                    s &= ", "
                End If
            Next
        End With
        s &= "}"
        If output_as = OutputAs.JObject Then
            r = JObject.Parse(s)
        ElseIf output_as = OutputAs.String_ Then
            r = s
        End If
        Return r
    End Function
#End Region

#Region "Web...?"

#Region "Text Formating"
    Private Shared Function TransformWord(word As Object, casing As TextCase)
        If CType(word, String).Length < 1 Then Return ""
        Dim s As String = CStr(word)
        Dim r = ""
        Select Case casing
            Case TextCase.Capitalize
                r = Mid(s, 1, 1).ToUpper & Mid(s, 2).ToLower
            Case TextCase.LowerCase
                r = s.ToLower
            Case TextCase.UpperCase
                r = s.ToUpper
            Case TextCase.None
                r = s
        End Select
        Return r
    End Function
    Private Shared Function TransformSingleLineText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        If CType(text, String).Trim.Length < 1 Then Return ""

        Dim d As List(Of String) = SplitTextInSplits(CStr(text), separator_, SideToReturn.AsListOfString)
        Dim o As New List(Of String)
        With d
            For i = 0 To .Count - 1
                o.Add(TransformWord(d(i), casing))
            Next
        End With
        Dim r = ""
        With o
            For i = 0 To .Count - 1
                r &= o(i)
                If i <> .Count - 1 Then r &= separator_
            Next
        End With
        Return r.Trim
    End Function

    Private Shared Function TransformMultiLineText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        Dim s As String = CStr(text).Trim
        Dim o As List(Of String) = SplitTextInSplits(s, vbCrLf, SideToReturn.AsListOfString)
        Dim f As New List(Of String)
        With o
            For i = 0 To .Count - 1
                f.Add(TransformSingleLineText(o(i).Trim, casing, separator_))
            Next
        End With
        Return ListToString(f)
    End Function
    Public Shared Function TransformText(text As Object, Optional casing As TextCase = TextCase.Capitalize, Optional separator_ As String = " ")
        Dim s As String = CStr(text)
        If s.Length < 1 Then Return ""
        If s.Contains(vbCrLf) Then
            Return TransformMultiLineText(text, casing, separator_)
        Else
            Return TransformSingleLineText(text, casing, separator_)
        End If
    End Function


#End Region

#Region "Main"
    ''' <summary>
    ''' Checks if a string is valid email address.
    ''' </summary>
    ''' <param name="email_">String to check</param>
    ''' <returns>True or False</returns>
    Public Shared Function IsEmail(email_ As String) As Boolean
        Dim isValid As Boolean = True
        If Not Regex.IsMatch(email_,
            "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$") Then
            isValid = False
        End If
        Return isValid
    End Function
    Public Shared Function firstWord(text As String) As String
        If text.Length < 1 Then Return ""

        Return SplitTextInTwo(text.Trim, " ", SideToReturn.Left)
    End Function

    Public Shared Function otherWords(text As String) As String
        If text.Length < 1 Then Return ""

        Return SplitTextInTwo(text.Trim, " ", SideToReturn.Right)
    End Function

    Public Shared Function lastThreeLetters(text As String) As Array
        Return {text.Chars(text.Length - 3), text.Chars(text.Length - 2), text.Chars(text.Length - 1)}
    End Function

    Public Shared Function IsVowel(text As String) As Boolean
        If text.Length < 1 Then Return False
        If text.ToLower = "a" Or
                 text.ToLower = "e" Or
                 text.ToLower = "i" Or
                 text.ToLower = "o" Or
                 text.ToLower = "u" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function IsConsonant(text As String) As Boolean
        If text.Length < 1 Then Return False
        If text.ToLower = "b" Or
                 text.ToLower = "c" Or
                 text.ToLower = "d" Or
                 text.ToLower = "f" Or
                 text.ToLower = "g" Or
                 text.ToLower = "h" Or
                 text.ToLower = "j" Or
                 text.ToLower = "k" Or
                 text.ToLower = "l" Or
                 text.ToLower = "m" Or
                 text.ToLower = "n" Or
                 text.ToLower = "p" Or
                 text.ToLower = "q" Or
                 text.ToLower = "r" Or
                 text.ToLower = "s" Or
                 text.ToLower = "t" Or
                 text.ToLower = "v" Or
                 text.ToLower = "w" Or
                 text.ToLower = "x" Or
                 text.ToLower = "y" Or
                 text.ToLower = "z" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function IsAlphabet(text As String) As Boolean
        If IsConsonant(text) Or IsVowel(text) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Turns statement to continuous tense. Works for ~80% scenarios.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="suffx"></param>
    ''' <returns></returns>
    Public Shared Function ToContinuous(text As String, Optional suffx As String = "") As String
        If text.Trim.Length < 1 Then Return ""
        Dim lastThree As Array = lastThreeLetters(text)
        If IsAlphabet(lastThree(0)) = False Or
                IsAlphabet(lastThree(1)) = False Or
                IsAlphabet(lastThree(2)) = False Then
            Return ""
        End If

        Dim a = lastThree(0).ToString.ToLower, b = lastThree(1).ToString.ToLower, c = lastThree(2).ToString.ToLower
        Dim prefx = ""

        If a = "i" And b = "n" And c = "g" Then
            Return text
        End If

        If IsConsonant(a) And IsVowel(b) And IsConsonant(c) Then
            prefx = text & Mid(text.Trim, text.Length, 1).Trim & "ing"
        ElseIf b = "i" And c = "e" Then
            prefx = Mid(text.Trim, 1, text.Length - 2).Trim & "ying"
        ElseIf IsVowel(a) And IsConsonant(b) And c = "e" Then
            prefx = Mid(text.Trim, 1, text.Length - 1).Trim & "ing"
        Else
            prefx = text.Trim & "ing"
        End If

        Return RTrim(prefx) & " " & LTrim(suffx)
    End Function
    Public Shared Function IsPhraseOrSentence(text As String) As Boolean
        If text.Trim.Length < 1 Then Return False
        If firstWord(text).Length > 0 And otherWords(text).Length > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Private Shared Function IsNumber(char_) As Boolean
    '	If CStr(char_).Length < 1 Then Return False
    '	Dim c As String = CStr(char_)
    '	For i = 0 To c.Length - 1
    '		If Mid(c, i, 1) = "1" Or
    '				Mid(c, i, 1) = "2" Or
    '				Mid(c, i, 1) = "3" Or
    '				Mid(c, i, 1) = "4" Or
    '				Mid(c, i, 1) = "5" Or
    '				Mid(c, i, 1) = "6" Or
    '				Mid(c, i, 1) = "7" Or
    '				Mid(c, i, 1) = "8" Or
    '				Mid(c, i, 1) = "9" Or
    '				Mid(c, i, 1) = "0" Then
    '		Else
    '			Return False
    '		End If
    '	Next
    '	Return True
    'End Function

    'Private Shared Function IsAlphabet(char_) As Boolean

    '	If CStr(char_).Length < 1 Then Return False
    '	Dim c As String = CStr(char_)
    '	For i = 0 To c.Length - 1
    '		If Mid(c, i, 1).ToLower = "a".ToLower Or
    '				Mid(c, i, 1).ToLower = "a".ToLower Or
    '				Mid(c, i, 1).ToLower = "b".ToLower Or
    '				Mid(c, i, 1).ToLower = "c".ToLower Or
    '				Mid(c, i, 1).ToLower = "d".ToLower Or
    '				Mid(c, i, 1).ToLower = "e".ToLower Or
    '				Mid(c, i, 1).ToLower = "f".ToLower Or
    '				Mid(c, i, 1).ToLower = "g".ToLower Or
    '				Mid(c, i, 1).ToLower = "h".ToLower Or
    '				Mid(c, i, 1).ToLower = "i".ToLower Or
    '				Mid(c, i, 1).ToLower = "j".ToLower Or
    '				Mid(c, i, 1).ToLower = "k".ToLower Or
    '				Mid(c, i, 1).ToLower = "l".ToLower Or
    '				Mid(c, i, 1).ToLower = "m".ToLower Or
    '				Mid(c, i, 1).ToLower = "n".ToLower Or
    '				Mid(c, i, 1).ToLower = "o".ToLower Or
    '				Mid(c, i, 1).ToLower = "p".ToLower Or
    '				Mid(c, i, 1).ToLower = "q".ToLower Or
    '				Mid(c, i, 1).ToLower = "r".ToLower Or
    '				Mid(c, i, 1).ToLower = "s".ToLower Or
    '				Mid(c, i, 1).ToLower = "t".ToLower Or
    '				Mid(c, i, 1).ToLower = "u".ToLower Or
    '				Mid(c, i, 1).ToLower = "v".ToLower Or
    '				Mid(c, i, 1).ToLower = "w".ToLower Or
    '				Mid(c, i, 1).ToLower = "x".ToLower Or
    '				Mid(c, i, 1).ToLower = "y".ToLower Or
    '				Mid(c, i, 1).ToLower = "z".ToLower Or
    '				Mid(c, i, 1).ToLower = "_".ToLower Then
    '		Else
    '			Return False
    '		End If
    '	Next
    '	Return True
    'End Function
    Public Shared Function MakeValid(str As String, valid_as As ValidAs)
        If str.Length < 1 Then Return ""
        Dim r As String = ""
        With str
            If valid_as = ValidAs.FileName Then
                For i = 0 To .Length - 1
                    If IsAlphabet(Mid(str, i, 1)) Or IsNumeric(Mid(str, i, 1)) Then
                        r &= Mid(str, i, 1)
                    End If
                Next
            ElseIf valid_as = ValidAs.FilePath_InJSON Then
                r = str.Trim.Replace("\", "\\")
            ElseIf valid_as = ValidAs.FilePath_Regular Then
                r = str.Trim.Replace("\\", "\")
            End If
        End With
        Return r.Trim
    End Function
#End Region

#Region "Map"
    Public Function MapProvidersList() As List(Of String)
        Dim l_ As New List(Of String)

        l_.Add("Bing Hybrid")
        l_.Add("Bing")

        l_.Add("ArcGIS StreetMap World 2D")
        l_.Add("ArcGIS World Street")
        l_.Add("Google China Hybrid")
        l_.Add("Google China")
        l_.Add("Google Hybrid")
        l_.Add("Google")
        l_.Add("OpenCycle Landscape")
        l_.Add("OpenCycle Transport")
        l_.Add("OpenCycle")
        l_.Add("WikiMapia")
        l_.Add("Yandex Hybrid")
        l_.Add("Yandex")

        Return l_
    End Function

#End Region

#Region "Support Functions"


    Public Shared Function FullNameFromNames(FirstName As String, LastName As String, Optional TitleOfCourtesy As String = "", Optional MiddleName As String = "", Optional last_name_first As Boolean = False, Optional include_title_of_courtesy As Boolean = True, Optional separate_first_or_last_name_with_comma As Boolean = False, Optional include_middle_name As Boolean = False) As String
        Dim r As String = ""
        If last_name_first Then
            If include_middle_name = True And MiddleName.Trim.Length > 0 Then
                r = LastName.Trim & " " & FirstName.Trim & " " & MiddleName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = LastName.Trim & ", " & FirstName.Trim & " " & MiddleName.Trim
                End If
            Else
                r = LastName.Trim & " " & FirstName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = LastName.Trim & ", " & FirstName.Trim
                End If
            End If
        Else
            r = FirstName.Trim & " " & LastName.Trim
            If separate_first_or_last_name_with_comma Then
                r = FirstName.Trim & ", " & LastName.Trim
            End If
            If include_middle_name = True And MiddleName.Trim.Length > 0 Then
                r = FirstName.Trim & " " & MiddleName & " " & LastName.Trim
                If separate_first_or_last_name_with_comma Then
                    r = FirstName.Trim & ", " & MiddleName & " " & LastName.Trim
                End If
            End If
        End If
        If include_title_of_courtesy And TitleOfCourtesy.Trim.Length > 0 Then
            r = TitleOfCourtesy & " " & r
        End If
        Return r
    End Function

    Public Shared Function ContainsAlphabet(str) As Boolean
        Dim r As Boolean = False
        If CType(str, String).ToLower.Contains("a") Or
            CType(str, String).ToLower.Contains("b") Or
            CType(str, String).ToLower.Contains("c") Or
            CType(str, String).ToLower.Contains("d") Or
            CType(str, String).ToLower.Contains("e") Or
            CType(str, String).ToLower.Contains("f") Or
            CType(str, String).ToLower.Contains("g") Or
            CType(str, String).ToLower.Contains("h") Or
            CType(str, String).ToLower.Contains("i") Or
            CType(str, String).ToLower.Contains("j") Or
            CType(str, String).ToLower.Contains("k") Or
            CType(str, String).ToLower.Contains("l") Or
            CType(str, String).ToLower.Contains("m") Or
            CType(str, String).ToLower.Contains("n") Or
            CType(str, String).ToLower.Contains("o") Or
            CType(str, String).ToLower.Contains("p") Or
            CType(str, String).ToLower.Contains("q") Or
            CType(str, String).ToLower.Contains("r") Or
            CType(str, String).ToLower.Contains("s") Or
            CType(str, String).ToLower.Contains("t") Or
            CType(str, String).ToLower.Contains("u") Or
            CType(str, String).ToLower.Contains("v") Or
            CType(str, String).ToLower.Contains("w") Or
            CType(str, String).ToLower.Contains("x") Or
            CType(str, String).ToLower.Contains("y") Or
            CType(str, String).ToLower.Contains("z") Or
            CType(str, String).ToLower.Contains("-") Then
            r = True
        End If

        If str.ToString.Length < 1 Then r = False

        Return r
    End Function

    Public Shared Function ContainsNumber(str) As Boolean
        Return Val(str) <> 0
    End Function

    Public Shared Function GetEnum(instance_of_enum As Object, Optional values_instead As Boolean = False) As List(Of String)
        Dim l As New List(Of String)

        If values_instead Then
            For Each i As Integer In [Enum].GetValues(instance_of_enum.GetType)
                l.Add(i)
            Next
            Return l
        End If

        For Each i In instance_of_enum.GetType().GetEnumNames()
            l.Add(i)
        Next
        Return l

    End Function

    Public Shared Function DropToEnum(str As String) As String
        Return str.Replace(" ", "_")
    End Function
    Public Shared Function DropToEnum(str As String, replacement As String) As String
        Return str.Replace(" ", replacement)
    End Function

    Public Shared Function EnumToDrop(str As String) As String
        Return str.Replace("_", " ")
    End Function
    Public Shared Function EnumToDrop(str As String, replacement As String) As String
        Return str.Replace("_", replacement)
    End Function

    'Public Shared Function stringToEnum(str As String) As String
    '    Return DropToEnum(str)
    'End Function
    'Public Shared Function stringToEnum(str As String, replacement As String) As String
    '    Return DropToEnum(str, replacement)
    'End Function

    Public Shared Function EnumToString(str As String) As String
        Return EnumToDrop(str)
    End Function
    Public Shared Function EnumToString(str As String, replacement As String) As String
        Return EnumToDrop(str, replacement)
    End Function
    Public Shared Function TextToEnum(str As String) As String
        Return DropToEnum(str)
    End Function
    Public Shared Function TextToEnum(str As String, replacement As String) As String
        Return DropToEnum(str, replacement)
    End Function

    Public Shared Function EnumToText(str As String) As String
        Return EnumToDrop(str)
    End Function
    Public Shared Function EnumToText(str As String, replacement As String) As String
        Return EnumToDrop(str, replacement)
    End Function

    Public Shared Function Percentage(score_, max_score)
        Return (score_ / max_score) * 100
    End Function
    Public Shared Function Grade(score_, max_score)
        Dim g_s = Percentage(score_, max_score)
        Dim r_s = ""

        If g_s >= 70 Then r_s = "A"

        If g_s >= 60 And g_s < 70 Then r_s = "B"

        If g_s >= 50 And g_s < 60 Then r_s = "C"

        If g_s >= 40 And g_s < 50 Then r_s = "D"

        If g_s >= 30 And g_s < 40 Then r_s = "E"

        If g_s < 30 Then r_s = "F"

        Return r_s

    End Function

    Public Shared Function GradeAfterPercentage(percentage_)
        Dim g_s = percentage_
        Dim r_s = ""

        If g_s >= 70 Then r_s = "A"

        If g_s >= 60 And g_s < 70 Then r_s = "B"

        If g_s >= 50 And g_s < 60 Then r_s = "C"

        If g_s >= 40 And g_s < 50 Then r_s = "D"

        If g_s >= 30 And g_s < 40 Then r_s = "E"

        If g_s < 30 Then r_s = "F"

        Return r_s

    End Function
    Public Shared Function Gross(quantity_bought As String, discount_units As String, unit_price As String, discount As String, use_discount As Boolean)
        quantity_bought = (quantity_bought)
        discount_units = (discount_units)
        unit_price = (unit_price)
        discount = (discount)
        Dim return_value
        If use_discount = False Or quantity_bought < (discount_units + 1) Then
            return_value = quantity_bought * unit_price
            GoTo 5
        End If
        'convert discount to decimal
        Dim discount_decimal = discount / 100
        'how much of price is off
        Dim how_much_off = discount_decimal * unit_price
        'what is the price now
        Dim new_price = unit_price - how_much_off
        'how many items uses the price now, get gross now
        Dim items_to_use_discount = (quantity_bought - discount_units) + 1
        Dim gross_with_discount = items_to_use_discount * new_price
        'how many items uses the old price, get gross then
        Dim items_to_do_without_discount = quantity_bought - items_to_use_discount
        Dim gross_without_discount = items_to_do_without_discount * unit_price
        'return gross now + gross then
        return_value = gross_without_discount + gross_with_discount
5:
        Return return_value
    End Function

    Public Shared Function Spaces(stringToAdd As String, maximum_length As Integer, Optional EndPoint As Boolean = True) As String
        Dim current_length As String = stringToAdd.Length
        Dim trail_ As String = ""
        Dim return_trail As String = ""
        If current_length < maximum_length Then
            For s__% = 1 To (maximum_length - current_length)
                return_trail &= " "
            Next
            GoTo 3
        ElseIf current_length >= maximum_length Then
            return_trail = ""
        End If

2:
        '		If EndPoint = False Then
        '			If stringToAdd.Length > (maximum_length + 1) Then
        '				stringToAdd = CInt((stringToAdd) - 1) & "+"
        '			End If
        '		End If
3:
        Return stringToAdd & return_trail
    End Function

#End Region

#Region "Internal"

    Public Shared Function DictionaryKeys(dict As Dictionary(Of Integer, String), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of Integer, String) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Keys(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function
    Public Shared Function DictionaryValues(dict As Dictionary(Of Integer, String), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of Integer, String) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Values(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function
    Public Shared Function DictionaryContains(dict As Dictionary(Of Integer, String), str_sought As String) As Boolean
        Dim l As List(Of String) = DictionaryValues(dict)
        Dim l_temp As New List(Of String)
        For i = 0 To l.Count - 1
            l_temp.Add(l(i).ToLower)
        Next
        Return l_temp.Contains(str_sought.ToLower)
    End Function
    Public Shared Function DictionaryKey(str_sought As String, dict As Dictionary(Of Integer, String)) As String
        Dim key As Integer = Nothing
        For Each k As String In dict.Keys
            If dict.Item(k) = str_sought Then
                key = k
            End If
        Next
        Return key
    End Function
    Public Shared Function DictionaryValue(dict As Dictionary(Of Integer, String), key As Integer) As String
        Return dict.Item(key)
    End Function



    Public Shared Function DictionaryKeys(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of String, Object) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Keys(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function
    Public Shared Function DictionaryValues(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.AsListOfString)
        Dim d As Dictionary(Of String, Object) = dict
        Dim l As New List(Of String)
        With d
            For i = 0 To .Count - 1
                l.Add(d.Values(i))
            Next
        End With
        If side_to_return = SideToReturn.AsListOfString Then
            Return l
        ElseIf side_to_return = SideToReturn.AsArray Then
            Return l.ToArray
        End If
    End Function

    Public Shared Function DictionaryContains(dict As Dictionary(Of String, Object), str_sought As String) As Boolean
        Dim l As List(Of String) = DictionaryValues(dict)
        Dim l_temp As New List(Of String)
        For i = 0 To l.Count - 1
            l_temp.Add(l(i).ToLower)
        Next
        Return l_temp.Contains(str_sought.ToLower)
    End Function

    Public Shared Function DictionaryKey(str_sought As String, dict As Dictionary(Of String, Object)) As String
        Dim key As Integer = Nothing
        For Each k As String In dict.Keys
            If dict.Item(k) = str_sought Then
                key = k
            End If
        Next
        Return key
    End Function
    Public Shared Function DictionaryValue(dict As Dictionary(Of String, Object), key As String) As String
        Return dict.Item(key)
    End Function

    Public Shared Function StringToDictionary(ByVal text As String, ByVal delimiter As String, ByVal equator As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)()
        Dim temp As List(Of String) = StringToList(text, delimiter)
        For i As Integer = 0 To temp.Count - 1
            Dim tokens() As String = temp(i).Split(equator)
            result.Add(tokens(0).Trim(), tokens(1).Trim())
        Next i
        Return result
    End Function

    Public Shared Function DictionaryToList(dict As Dictionary(Of String, Object), Optional side_to_return As SideToReturn = SideToReturn.Right) As IEnumerable(Of Object)

        Dim result = Nothing
        Select Case side_to_return
            Case SideToReturn.Values
                result = dict.Values.ToList()
            Case SideToReturn.Right
                result = dict.Values.ToList()
            Case SideToReturn.Keys
                result = dict.Keys.ToList()
            Case SideToReturn.Left
                result = dict.Keys.ToList()


        End Select
        Return result


        'Dim d As Dictionary(Of String, Object) = dict
        'Dim l As New List(Of String)
        'Dim r As New List(Of Object)
        'With d
        '    For i = 0 To .Count - 1
        '        If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Left Then
        '            l.Add(d.Keys(i))
        '        End If
        '        If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Right Then
        '            r.Add(d.Values(i))
        '        End If
        '    Next
        'End With
        'If side_to_return = SideToReturn.Left Then
        '    Return ListToString(l)
        'ElseIf side_to_return = SideToReturn.Right Then
        '    Return ListToString(r)
        'Else 'side_to_return = SideToReturn.AsArray
        '    Return {l, r}
        'End If
    End Function

    Public Shared Function DictionaryToList(dict As Dictionary(Of String, String), Optional side_to_return As SideToReturn = SideToReturn.Right) As IEnumerable(Of String)

        Dim result = Nothing
        Select Case side_to_return
            Case SideToReturn.Values
                result = dict.Values.ToList()
            Case SideToReturn.Right
                result = dict.Values.ToList()
            Case SideToReturn.Keys
                result = dict.Keys.ToList()
            Case SideToReturn.Left
                result = dict.Keys.ToList()


        End Select
        Return result


        'Dim d As Dictionary(Of String, Object) = dict
        'Dim l As New List(Of String)
        'Dim r As New List(Of Object)
        'With d
        '    For i = 0 To .Count - 1
        '        If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Left Then
        '            l.Add(d.Keys(i))
        '        End If
        '        If side_to_return = SideToReturn.AsArray Or side_to_return = SideToReturn.Right Then
        '            r.Add(d.Values(i))
        '        End If
        '    Next
        'End With
        'If side_to_return = SideToReturn.Left Then
        '    Return ListToString(l)
        'ElseIf side_to_return = SideToReturn.Right Then
        '    Return ListToString(r)
        'Else 'side_to_return = SideToReturn.AsArray
        '    Return {l, r}
        'End If
    End Function

    ''' <summary>
    ''' Joins list items to form a string. 
    ''' </summary>
    ''' <param name="list">List or ReadOnlyCollection</param>
    ''' <param name="delimiter"></param>
    ''' <param name="format_output"></param>
    ''' <returns></returns>

    Public Shared Function ListToString(list As Object, Optional delimiter As String = vbCrLf, Optional format_output As Boolean = False) As String
        Dim delimeter As String = delimiter
        Dim l 'As List(Of String)
        If TypeOf list Is ReadOnlyCollection(Of String) Then
            l = CType(list, ReadOnlyCollection(Of String))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Integer) Then
            l = CType(list, ReadOnlyCollection(Of Integer))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Decimal) Then
            l = CType(list, ReadOnlyCollection(Of Decimal))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Double) Then
            l = CType(list, ReadOnlyCollection(Of Double))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Short) Then
            l = CType(list, ReadOnlyCollection(Of Short))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Byte) Then
            l = CType(list, ReadOnlyCollection(Of Byte))
        ElseIf TypeOf list Is ReadOnlyCollection(Of Long) Then
            l = CType(list, ReadOnlyCollection(Of Long))
        Else
            l = CType(list, List(Of String))
        End If
        Dim r As String = ""
        With l
            For i As Integer = 0 To .Count - 1
                r &= l(i).ToString & delimeter
            Next
        End With
        If format_output Then
            Return PrepareForIO(r)
        Else
            Return r
        End If
    End Function

    'Public Shared Function ArrayToString(array As Array, Optional delimeter As String = vbCrLf, Optional format_output As Boolean = False) As String
    '    Dim r As String = ""
    '    With array
    '        For i As Integer = 0 To .Length - 1
    '            r &= array(i) & delimeter
    '        Next
    '    End With
    '    If format_output Then
    '        Return PrepareForIO(r)
    '    Else
    '        Return r
    '    End If

    'End Function


    Public Shared Function ArrayToList(array_ As Array, list_is As ListIsOf) As Object
        Dim a As Array = array_
        Dim l_string As New List(Of String)

        For Each i In a
            If list_is = ListIsOf.String_ Then
                l_string.Add(i)
            End If
        Next

        Return l_string
    End Function

    Public Shared Function StringToList(delimited_string As String, Optional delimeter As String = vbCrLf, Optional return_ As SideToReturn = SideToReturn.AsListOfString) As Object
        Return SplitTextInSplits(delimited_string, delimeter, return_)
    End Function

#End Region

#Region "DropText-Boolean"

    Public Enum DropTextPattern As Byte
        AlwaysNever
        YesNo
        OnOff
        OneZero
        TrueFalse
        IfNot
        SucceededFailed
        Cart
    End Enum

    Public Shared ReadOnly Property BooleanDropTextTrue As List(Of String) =
        {"yes, include this item", "always", "yes", "on", "one", "true", "if possible", "succeeded"}.ToList()
    Public Shared ReadOnly Property BooleanDropTextFalse As List(Of String) =
        {"no, remove this item", "never", "no", "off", "zero", "false", "not at all", "failed"}.ToList()
    Public Shared cart_inclusion_list As String() = {"Yes, Include This Item", "No, Remove This Item"}


    ''' <summary>
    ''' Converts boolean values to user-friendly pattern e.g. yes, never.
    ''' To convert to boolean, use DropTextToBoolean(str_ As String).
    ''' </summary>
    ''' <param name="boolean_val">True or False</param>
    ''' <param name="pattern_">always/never (default) OR a/n OR a, yes/no OR y/n OR y, on/off OR o/f OR o, 1/0, true/false OR t/f OR t, if/not OR i/n OR i</param>
    ''' <returns>Always/Never (default), Yes/No, On/Off, 1/0, True/Talse, If possible/Not at all</returns>
    Private Shared Function BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "Yes/No") As String

        Select Case Convert.ToBoolean(boolean_val)
            Case True
                If pattern_.ToLower = "" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "always/never" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "a/n" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "a" Then
                    Return "Always"
                End If
                If pattern_.ToLower = "yes/no" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "y/n" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "y" Then
                    Return "Yes"
                End If
                If pattern_.ToLower = "on/off" Then
                    Return "On"
                End If
                If pattern_.ToLower = "o/f" Then
                    Return "On"
                End If
                If pattern_.ToLower = "o" Then
                    Return "On"
                End If
                If pattern_.ToLower = "1/0" Then
                    Return "1"
                End If
                If pattern_.ToLower = "true/false" Then
                    Return "True"
                End If
                If pattern_.ToLower = "t/f" Then
                    Return "True"
                End If
                If pattern_.ToLower = "t" Then
                    Return "True"
                End If
                If pattern_.ToLower = "if/not" Then
                    Return "If possible"
                End If
                If pattern_.ToLower = "i/n" Then
                    Return "If possible"
                End If
                If pattern_.ToLower = "i" Then
                    Return "If possible"
                End If

            Case False
                If pattern_.ToLower = "" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "always/never" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "a/n" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "a" Then
                    Return "Never"
                End If
                If pattern_.ToLower = "yes/no" Then
                    Return "No"
                End If
                If pattern_.ToLower = "y/n" Then
                    Return "No"
                End If
                If pattern_.ToLower = "y" Then
                    Return "No"
                End If
                If pattern_.ToLower = "on/off" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "o/f" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "o" Then
                    Return "Off"
                End If
                If pattern_.ToLower = "1/0" Then
                    Return "0"
                End If
                If pattern_.ToLower = "true/false" Then
                    Return "False"
                End If
                If pattern_.ToLower = "t/f" Then
                    Return "False"
                End If
                If pattern_.ToLower = "t" Then
                    Return "False"
                End If
                If pattern_.ToLower = "if/not" Then
                    Return "Not at all"
                End If
                If pattern_.ToLower = "i/n" Then
                    Return "Not at all"
                End If
                If pattern_.ToLower = "i" Then
                    Return "Not at all"
                End If
        End Select
    End Function

    Public Shared Function BooleanToDropText(boolean_val As Boolean, Optional pattern_ As DropTextPattern = DropTextPattern.YesNo) As String

        Select Case Convert.ToBoolean(boolean_val)
            Case True
                If pattern_.ToString.ToLower = DropTextPattern.AlwaysNever.ToString.ToLower Then
                    Return "Always"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.YesNo.ToString.ToLower Then
                    Return "Yes"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OnOff.ToString.ToLower Then
                    Return "On"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OneZero.ToString.ToLower Then
                    Return "One"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.TrueFalse.ToString.ToLower Then
                    Return "True"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.IfNot.ToString.ToLower Then
                    Return "If possible"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.SucceededFailed.ToString.ToLower Then
                    Return "Succeeded"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.Cart.ToString.ToLower Then
                    Return "Yes, Include This Item"
                End If
            Case False
                If pattern_.ToString.ToLower = DropTextPattern.AlwaysNever.ToString.ToLower Then
                    Return "Never"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.YesNo.ToString.ToLower Then
                    Return "No"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OnOff.ToString.ToLower Then
                    Return "Off"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.OneZero.ToString.ToLower Then
                    Return "Zero"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.TrueFalse.ToString.ToLower Then
                    Return "False"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.IfNot.ToString.ToLower Then
                    Return "Not at all"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.SucceededFailed.ToString.ToLower Then
                    Return "Failed"
                End If
                If pattern_.ToString.ToLower = DropTextPattern.Cart.ToString.ToLower Then
                    Return "No, Remove This Item"
                End If
        End Select
    End Function
    ''' <summary>
    ''' Converts user-friendly word like yes or always to boolean i.e. true or false.
    ''' To convert back to user-friendly word, use BooleanToDropText(boolean_val As Boolean, Optional pattern_ As String = "always/never")
    ''' </summary>
    ''' <example>
    ''' <code>
    ''' Dim userDecision as string = DropTextToBoolean(dropText.text)
    ''' </code>
    ''' </example>
    ''' <param name="str_">always/never (default), yes/no, on/off, 1/0, true/false, if/not</param>
    ''' <returns>True or False</returns>
    Public Shared Function DropTextToBoolean(str_ As String) As Boolean
        If BooleanDropTextTrue.Contains(str_.ToLower) Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region



#End Region

#Region "Support"
    Public Shared Function FileTypeToString(fileType As FileType) As String
        If fileType = FileType.any Then
            Return "*.*"
        Else
            Return "*." & fileType.ToString.Replace("_", "")
        End If
        'fileType.repl
        'Select Case fileType
        '    Case FileType._3gp
        '        Return "*.3gp"
        '    Case FileType._3gpp
        '        Return "*.3gpp"
        '    Case FileType._3pg2
        '        Return "*.3pg2"
        '    Case FileType.any
        '        Return "*.*"
        '        'Case filetyp
        '    Case Else
        '        Return "*." & fileType.ToString
        'End Select
    End Function


    Private Shared Function CustomMarkup(str_ As String) As String
        Return str_
    End Function

    Public Shared Function PrepareForIO(str_ As String, Optional output_ As FormatFor = FormatFor.Web) As String
        If output_ = FormatFor.Custom Then Return CustomMarkup(str_)

        Dim trimmed_, CR_less, CRLFless, TABless As String
        trimmed_ = str_.Trim
        str_ = trimmed_

        If output_ = FormatFor.Web Then
            CRLFless = str_.Replace(vbCrLf, "<br />")
            str_ = CRLFless
        ElseIf output_ = FormatFor.JavaScript Then
            CRLFless = str_.Replace(vbCrLf, "\n")
            str_ = CRLFless
        End If

        If output_ = FormatFor.Web Then
            CR_less = str_.Replace(vbCrLf, "<br />")
            str_ = CR_less
        ElseIf output_ = FormatFor.JavaScript Then
            CR_less = str_.Replace(vbCrLf, "\n")
            str_ = CR_less
        End If

        If output_ = FormatFor.Web Then
            TABless = str_.Replace(vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
            str_ = TABless
        ElseIf output_ = FormatFor.JavaScript Then
            TABless = str_.Replace(vbTab, "\t")
            str_ = TABless
        End If

        Return str_
    End Function
    Private Shared Function WebPageMarkup(str_ As String, web_content_in As WrapWebContent, output_ As FormatFor) As String
        Dim r
        If web_content_in = WrapWebContent.AsIs Then
            r = ToIO(str_, output_)
        ElseIf web_content_in = WrapWebContent.Div Then
            r = "<div>" & ToIO(str_, output_) & "</div>"
        ElseIf web_content_in = WrapWebContent.P Then
            r = "<p>" & ToIO(str_, output_) & "</p>"
        End If

        Return r
    End Function


    Public Shared Function ToWeb(str_ As String, filename_ As String, server_ As String, username_ As String, password As String, Optional web_content_in As WrapWebContent = WrapWebContent.Div, Optional output_ As FormatFor = FormatFor.Web)
        Dim s As String = WebPageMarkup(str_, web_content_in, output_)
        'send to server
    End Function
    Public Shared Function ToWeb(str_ As String, filename_ As String, Optional web_content_in As WrapWebContent = WrapWebContent.Div, Optional output_ As FormatFor = FormatFor.Web)
        Dim s As String = WebPageMarkup(str_, web_content_in, output_)
        Try
            My.Computer.FileSystem.WriteAllText(filename_, s, False)
        Catch ex As Exception

        End Try
        Return s
    End Function
    'Public Shared Sub ToQR(str_ As String, filename_ As String, server_ As String, username_ As String, password As String)
    '    Dim bmp = QREncode(str_, filename_)
    '    'send to server
    'End Sub

    Public Shared Function ToIO(str_ As String, Optional output_ As FormatFor = FormatFor.Web) As String
        Return PrepareForIO(str_, output_)
    End Function
    Public Shared Function RemoveHTMLFromText(ByVal html_markup As String) As String
        Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("<[^>]*>")
        Return rx.Replace(html_markup, "")
    End Function

    Public Shared Function HTMLToText(ByVal html_markup As String) As String
        Return RemoveHTMLFromText(html_markup)
    End Function

    ''' <summary>
    ''' Adds a random number to the list, keeping all the list's items unique.
    ''' </summary>
    ''' <param name="random_inclusive_min">Where to start (used by random)</param>
    ''' <param name="random_exclusive_max">Where to stop (used by random)</param>
    ''' <param name="already_">The list that to add to</param>
    ''' <returns>The list with the new number</returns>

    Public Shared Function RandomList(random_inclusive_min As Integer, random_exclusive_max As Integer, already_ As List(Of Integer), Optional recycle_ As Boolean = True) As Integer
        Dim r_val
        If already_.Count >= random_exclusive_max Then
            If recycle_ = True Then
                already_.Clear()
            Else
                Return already_(already_.Count - 1)
            End If
        End If
2:
        Try
            r_val = Random_(random_inclusive_min, random_exclusive_max)
            If already_.Count > 0 And already_.Contains(r_val) Then
                GoTo 2
            Else
                already_.Add(r_val)
                Return r_val
            End If
        Catch
        End Try
    End Function

    ''' <summary>
    ''' Creates a shuffled (integer) list.
    ''' </summary>
    ''' <param name="items_from"></param>
    ''' <param name="items_to"></param>
    ''' <param name="items_"></param>
    ''' <returns></returns>
    Public Function CreateRandomList(items_from As Integer, items_to As Integer, items_ As List(Of Integer)) As List(Of Integer)
        For i As Integer = items_from To items_to
            RandomList(items_from, items_to + 1, items_)
        Next
        Return items_


        '		Dim temp_list As New List(Of Integer)
        '		g.CreateRandomList(1, 7, temp_list)

        '		'Then use the list, for example:
        '		If temp_list.Count < 1 Then Exit Sub
        '		With temp_list
        '			For j As Integer = 0 To .Count - 1
        '				u.Text &= .Item(j).ToString & vbCrLf
        '			Next
        '		End With

    End Function

    'Private Function CycleThrough(c As Control, Optional reverse_ As Boolean = False) As Integer
    '    Dim l As ListBox, d As ComboBox
    '    If TypeOf c Is ListBox Then
    '        l = c
    '        With l
    '            If .SelectedIndex < 1 Then .SelectedIndex = 0
    '            Try
    '                If reverse_ = True Then
    '                    .SelectedIndex -= 1
    '                Else
    '                    .SelectedIndex += 1
    '                End If
    '            Catch ex As Exception
    '                .SelectedIndex = 0
    '            End Try
    '        End With
    '        Return l.SelectedIndex
    '    End If
    '    If TypeOf c Is ComboBox Then
    '        d = c
    '        With d
    '            If .SelectedIndex < 1 Then .SelectedIndex = 0
    '            Try
    '                If reverse_ = True Then
    '                    .SelectedIndex -= 1
    '                Else
    '                    .SelectedIndex += 1
    '                End If
    '            Catch ex As Exception
    '                .SelectedIndex = 0
    '            End Try
    '        End With
    '        Return d.SelectedIndex
    '    End If
    'End Function

    Public Shared Function Random_(inclusive_min As Integer, exclusive_max As Integer) As Integer
        Dim generator As New Random
        Dim randomValue As Integer
        randomValue = generator.Next(inclusive_min, exclusive_max)
        Return randomValue
    End Function

    Public Shared Function LineBreak(times_ As Integer, Optional char_ As String = "*") As String
        Dim s As String = ""
        For i As Integer = 1 To times_
            s &= char_
        Next
        Return s
    End Function

    ''' <summary>
    ''' Returns the opposite of the (boolean) value.
    ''' </summary>
    ''' <param name="boolean_val">True or False</param>
    ''' <returns>The opposite of boolean_val</returns>
    Public Shared Function BooleanOpposite(boolean_val As Boolean) As Boolean
        Select Case boolean_val
            Case True
                Return False
            Case False
                Return True
        End Select
    End Function

    Private Shared Function gListToString(list_or_array As Object, Optional format_output As Boolean = False) As String
        Dim l As List(Of String)
        If TypeOf list_or_array Is Array Then
            Return gArrayToString(list_or_array)
        Else
            l = CType(list_or_array, List(Of String))
        End If
        Dim r As String = ""
        With l
            For i As Integer = 0 To .Count - 1
                r &= l(i) & vbCrLf
            Next
        End With
        If format_output Then
            Return PrepareForIO(r)
        Else
            Return r
        End If
    End Function
    Private Shared Function gArrayToString(array As Array, Optional format_output As Boolean = False) As String
        Dim r As String = ""
        With array
            For i As Integer = 0 To .Length - 1
                r &= array(i) & vbCrLf
            Next
        End With
        If format_output Then
            Return PrepareForIO(r)
        Else
            Return r
        End If

    End Function
    '   ''' <summary>
    '   ''' Splits text to left and right sides of separator_ and returns either left or right side depending on side_to_return.
    '   ''' </summary>
    '   ''' <param name="text_to_split"></param>
    '   ''' <param name="separator_"></param>
    '   ''' <param name="side_to_return"></param>
    '   ''' <returns>String</returns>
    '   Private Shared Function SplitTextInTwo(text_to_split As String, separator As String, side_to_return As SideToReturn) As String
    '	Dim r As Array = SplitText(text_to_split, separator)
    '	If side_to_return = SideToReturn.Left Then
    '		Return gListToString(r(0))
    '	Else
    '		Return gListToString(r(1))
    '	End If
    'End Function
    ''' <summary>
    ''' Prefer SplitTextInSplits over this...
    ''' Splits text to left and right sides of separator_ and returns either left or right side depending on side_to_return.
    ''' </summary>
    ''' <param name="string_to_split"></param>
    ''' <param name="separator">delimeter</param>
    ''' <param name="side_to_return">Left or Right</param>
    ''' <returns>String (left side), String (right side) or List(Of String) - left and right</returns>
    Public Shared Function SplitTextInTwo(string_to_split As String, separator As String, Optional side_to_return As SideToReturn = SideToReturn.Right)
        If string_to_split.Length < 1 Or separator.Length < 1 Then Return ""

        If side_to_return = SideToReturn.Left Then
            Return CType(string_to_split.Split(separator, 2, StringSplitOptions.None)(0), String)
        ElseIf side_to_return = SideToReturn.Right Then
            Try
                Return CType(string_to_split.Split(separator, 2, StringSplitOptions.None)(1), String)
            Catch ex As Exception
                Return ""
            End Try
        Else
            Try
                Dim r As New List(Of String)
                r.Add(CType(string_to_split.Split(separator, 2, StringSplitOptions.None)(0), String))
                r.Add(CType(string_to_split.Split(separator, 2, StringSplitOptions.None)(1), String))
                Return r
            Catch ex As Exception
                Return ""
            End Try
        End If

    End Function

    ''' <summary>
    ''' Splits string into multiple parts.
    ''' </summary>
    ''' <param name="string_to_split"></param>
    ''' <param name="separator"></param>
    ''' <param name="side_to_return"></param>
    ''' <returns></returns>
    Public Shared Function SplitTextInSplits(string_to_split As String, separator As String, Optional side_to_return As SideToReturn = SideToReturn.AsArray)
        If string_to_split.Length < 1 Then Return Nothing

        Dim str As String
        Dim strArr() As String
        str = string_to_split
        strArr = str.Split(separator)

        Dim l As New List(Of String)

        Select Case side_to_return
            Case SideToReturn.AsArray
                Return strArr
            Case SideToReturn.AsListOfString
                With strArr
                    For i = 0 To .Length - 1
                        l.Add(strArr(i))
                    Next
                End With
                Return l
            Case SideToReturn.AsListToString
                Return ListToString(strArr)
        End Select
    End Function
    'Private Shared Function number_of_splits(string_to_split As String, separator As String)
    '	Dim s As String = string_to_split.Trim
    '	Dim sep As String = separator
    '	Dim counter = 0
    '	For i = 1 To s.Length
    '		If Mid(s, i, sep.Length) = sep Then counter += 1
    '	Next
    '	Return counter
    'End Function
    'Public Shared Function OccurrencesInString(string_ As String, string_that_occur_in_string_ As String)
    '	Dim s As String = string_.Trim
    '	Dim sep As String = string_that_occur_in_string_
    '	Dim counter = 0
    '	For i = 1 To s.Length
    '		If Mid(s, i, sep.Length) = sep Then counter += 1
    '	Next
    '	Return counter
    'End Function

    '''' <summary>
    '''' Splits text to left and right sides of separator_. Pass the result as argument to Use NModule.NFunctions.ListToString to get the string, or use the overload that accepts SideToReturn instead.
    '''' </summary>
    '''' <param name="text_to_split"></param>
    '''' <param name="separator_"></param>
    '''' <returns>Array of left (list) and right (list)</returns>
    'Private Shared Function SplitText(text_to_split As String, separator_ As String) As Array
    '	Dim a As String = text_to_split
    '	Dim l As New List(Of String)
    '	Dim r As New List(Of String)
    '	If a.ToLower.Contains(separator_.ToLower) Then
    '		For j As Integer = 1 To a.Length
    '			If Mid(a, j, separator_.Length + 2).ToLower = " " & separator_.ToLower & " " Then
    '				Dim r_ As String = a.Substring(j).Replace("'", "")
    '				l.Add(a.Substring(0, j).Replace("'", "").Trim)
    '				r.Add(r_.Replace(separator_, "").Trim)
    '			End If
    '		Next
    '	Else
    '		l.Add(a)
    '		r.Add("")
    '	End If

    '	Return {l, r}
    'End Function
    '''' <summary>
    '''' Splits each text in List(of String) to left and right sides of separator_. Pass the items in the result individually as argument to Use NModule.NFunctions.ListToString to get the string.
    '''' </summary>
    '''' <param name="texts_to_split"></param>
    '''' <param name="separator_"></param>
    '''' <returns>Array of left (list) and right (list)</returns>
    'Public Shared Function SplitText(texts_to_split As List(Of String), separator_ As String) As Array
    '	Dim a As List(Of String) = texts_to_split
    '	Dim l As New List(Of String)
    '	Dim r As New List(Of String)
    '	For i As Integer = 0 To a.Count - 1
    '		If a(i).ToLower.Contains(separator_.ToLower) Then
    '			For j As Integer = 1 To a(i).Length
    '				If Mid(a(i), j, separator_.Length + 2).ToLower = " " & separator_.ToLower & " " Then
    '					Dim r_ As String = a(i).Substring(j).Replace("'", "")
    '					l.Add(a(i).Substring(0, j).Replace("'", "").Trim)
    '					r.Add(r_.Replace(separator_, "").Trim)
    '				End If
    '			Next
    '		Else
    '			l.Add(a(i))
    '			r.Add("")
    '		End If
    '	Next
    '	Return {l, r}
    'End Function

    Public Shared Function Stripped(str_ As String) As String
        Dim s As String = str_
        s = s.Replace(vbCr, " ").Trim
        s = s.Replace(vbCrLf, " ").Trim
        Return s.Trim
    End Function

    'Public Shared Function SplitString(string_ As String, separator_ As String) As Array
    '	separator_ = separator_.Trim
    '	Dim return_() As String = {}

    '	Dim string_parts() As String
    '	Dim left_ As String
    '	Dim right_ As String
    '	string_parts = string_.Split(separator_.ToCharArray, 2)

    '	If string_parts.Length = 2 Then
    '		left_ = string_parts(0)
    '		right_ = string_parts(1)
    '		If left_.Length > 0 And right_.Length > 0 Then
    '			return_ = {left_, right_}
    '		Else
    '			return_ = {Val(string_)}
    '		End If
    '	Else
    '		return_ = {Val(string_)}
    '	End If

    '	Return return_
    'End Function

    Public Shared Function NewID(Optional case_acct_or_date_time As String = "date_time", Optional replace_guid As String = "") As String
        '        Dim m As New FormatWindow
        Dim raw_id As String = System.Guid.NewGuid().ToString, counter As Integer

        If replace_guid.Trim.Length > 0 Then
            raw_id = replace_guid.Trim
            GoTo 2
        End If

        For i% = 1 To raw_id.Length
            If Mid(raw_id, i, 1) = "-" Then
                counter += 1
                If counter = 2 Then
                    raw_id = Mid(raw_id, 1, i - 1)
                ElseIf counter = 0 Then
                    Exit For
                End If
            End If
        Next

2:
        Select Case case_acct_or_date_time.ToLower
            Case "acct"
                raw_id = raw_id & "-" & My.Computer.Clock.LocalTime.Date.ToShortDateString
            Case "date_time"
                raw_id = raw_id & "-" & Now.Year.ToString & "." & Now.Month.ToString & "." & Now.Day.ToString & "-" & Now.Hour.ToString & "." & LeadingZero(Now.Minute.ToString)
            Case ""
        End Select
        Return raw_id
    End Function
    Public Shared Function NewID(IDPattern_ As IDPattern, Optional prefx As String = "") As String
        Select Case IDPattern_
            Case IDPattern.Long_
                Return NewGUID(prefx)
            Case IDPattern.Long_DateTime
                Return NewGUID(prefx, True)
            Case IDPattern.Short_
                Return NewID("", prefx)
            Case IDPattern.Short_DateOnly
                Return NewID("acct")
            Case IDPattern.Short_DateTime
                Return NewID("date_time")
        End Select
    End Function
    Public Shared Function NewGUID(IDPattern_ As IDPattern, Optional prefx As String = "") As String
        Select Case IDPattern_
            Case IDPattern.Long_
                Return NewGUID(prefx)
            Case IDPattern.Long_DateTime
                Return NewGUID(prefx, True)
            Case IDPattern.Short_
                Return NewID("", prefx)
            Case IDPattern.Short_DateOnly
                Return NewID("acct")
            Case IDPattern.Short_DateTime
                Return NewID("date_time")
        End Select
    End Function
    Public Shared Function NewGUIDr(Optional IDPattern_ As IDPattern = IDPattern.Short_, Optional prefx As String = "") As String
        Dim r As String
        Select Case IDPattern_
            Case IDPattern.Long_
                r = NewGUID(prefx)
            Case IDPattern.Long_DateTime
                r = NewGUID(prefx, True)
            Case IDPattern.Short_
                r = NewID("", prefx)
            Case IDPattern.Short_DateOnly
                r = NewID("acct")
            Case IDPattern.Short_DateTime
                r = NewID("date_time")
        End Select
        Dim a As String = r.Replace("-", "")
        Return a.Replace("_", "")

    End Function

    Public Shared Function NewGUID(Optional prefx As String = "", Optional suffx As Boolean = False) As String
        Dim return__ As String = ""
        If prefx.Length > 0 Then return__ = prefx & "-"
        If suffx = True Then return__ &= Now.ToShortDateString & "-" & Now.ToLongTimeString & "-"
        return__ = return__.Replace(" ", "")
        Dim r As String = return__.Replace("/", "-")
        Dim r_ As String = r.Replace(":", "-")
        Dim r3 As String = r_.Replace("AM", "-am")
        Dim return_ As String = r3.Replace("PM", "-pm")

        If return_.Length > 0 Then
            Return return_ & System.Guid.NewGuid().ToString
        Else
            Return System.Guid.NewGuid().ToString
        End If
    End Function

    Public Shared Function LeadingZero(number As String) As String
        If number.Length < 2 Then Return "0" & number Else : Return number
    End Function

    Public Shared Function GetTimeZone() As String
        Dim offset_ As String = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString
        Dim suffix_ As String = " (UTC"
        If InStr(offset_, "-") = False Then suffix_ = " (UTC+"
        Return TimeZone.CurrentTimeZone.StandardName & suffix_ & TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now).ToString & ")"
    End Function

    Private Shared Function TimeZone_() As String
        Return GetTimeZone()
    End Function

    Public Shared Function FullTimeWithZone() As String
        Return Now.ToLongTimeString & "  " & GetTimeZone()
    End Function

    Public Shared Function FullDateWithZone() As String
        Return Now.ToShortDateString & "  " & GetTimeZone()
    End Function

    Public Shared Function FullDateAsLongWithZone() As String
        Return Now.ToLongDateString & "  " & GetTimeZone()
    End Function

    Public Shared Function ExplicitDate(Optional time_ As String = "", Optional date_ As String = "") As String
        Return FullTimeAsLong(time_, date_) & ",  " & GetTimeZone()
    End Function

    Public Shared Function FullDate(Optional time_ As String = "", Optional date_ As String = "") As String
        'same as FullTime, but puts date first
        Dim val As String = ""
        If time_.Length < 1 Then time_ = Now.ToLongTimeString

        If date_.Length < 1 Then date_ = Now.ToShortDateString

        Return date_ & ",  " & time_
    End Function

    Public Shared Function FullTime(Optional time_ As String = "", Optional date_ As String = "") As String
        'same as FullDate, but puts time first
        Dim val As String = ""
        If time_.Length < 1 Then time_ = Now.ToLongTimeString

        If date_.Length < 1 Then date_ = Now.ToShortDateString

        Return time_ & ",  " & date_
    End Function

    ''' <summary>
    ''' Returns specified date and time (or current) in Long Date and Long Time, time first. Equivalent to FullTimeAsLong.
    ''' </summary>
    ''' <param name="time_">The time in any format</param>
    ''' <param name="date_">The date in any format</param>
    ''' <returns>Date and time</returns>
    Public Shared Function FullDateAsLong(Optional time_ As String = "", Optional date_ As String = "") As String
        Dim val As String = ""
        If time_.Length < 1 Then
            time_ = Now.ToLongTimeString
        Else
            time_ = DateTime.Parse(time_).ToLongTimeString
        End If

        If date_.Length < 1 Then
            date_ = Now.ToLongDateString
        Else
            date_ = DateTime.Parse(date_).ToLongDateString
        End If

        Return date_ & ",  " & time_
    End Function

    ''' <summary>
    ''' Returns specified date and time (or current) in Long Date and Long Time, time first. Equivalent to FullDateAsLong.
    ''' </summary>
    ''' <param name="time_">The time in any format</param>
    ''' <param name="date_">The date in any format</param>
    ''' <returns>Date and time</returns>
    Public Shared Function FullTimeAsLong(Optional time_ As String = "", Optional date_ As String = "") As String
        Dim val As String = ""
        If time_.Length < 1 Then
            time_ = Now.ToLongTimeString
        Else
            time_ = DateTime.Parse(time_).ToLongTimeString
        End If

        If date_.Length < 1 Then
            date_ = Now.ToLongDateString
        Else
            date_ = DateTime.Parse(date_).ToLongDateString
        End If

        Return time_ & ",  " & date_
    End Function

    Public Shared Function DateToShort(obj_ As Object, Optional convert_date_to_short As Boolean = True) As String
        Try
            If TypeOf obj_ Is Date Or TypeOf obj_ Is DateTime Or IsDate(obj_) Then
                If convert_date_to_short Then
                    Try
                        Return Date.Parse(Date.Parse(obj_).ToShortDateString())
                    Catch
                    End Try
                Else
                    Try
                        Return obj_.ToString
                    Catch
                    End Try
                End If
            Else
                Try
                    Return obj_.ToString
                Catch
                End Try
            End If
        Catch
        End Try
    End Function
    Public Shared Function DateToShort_(obj_ As Object, Optional convert_date_to_short As Boolean = True, Optional use_short_format As Boolean = True) As String
        If TypeOf obj_ Is Date Or TypeOf obj_ Is DateTime Or IsDate(obj_) Then
            If convert_date_to_short Then
                If use_short_format Then
                    Return Date.Parse(Date.Parse(obj_).ToShortDateString())
                Else
                    Return Date.Parse(Date.Parse(obj_).ToLongDateString())
                End If
            End If
        Else
            Return obj_.ToString
        End If
    End Function

    Public Shared Function FloorValue(val_ As String) As String
        Dim s As Integer
        For i As Integer = 1 To val_.Length
            If Mid(val_, i, 1) = "." Then
                s = i - 1
                Exit For
            End If
        Next
        Return Mid(val_, 1, s)
    End Function

    Public Shared Function RoundNumber(val_)
        Dim val__
        If Val(val_) = 0 Then
            val__ = 0
        Else
            val__ = val_
        End If
        Dim return_ = FormatNumber(val__, 2, TriState.False, TriState.False, TriState.False)
        If return_.ToString = ".00" Then return_ = "0.00"
        Return return_
    End Function
    Public Shared Function CalculateSince(date_ As Date, Optional interval_ As DateInterval = DateInterval.Day, Optional suffixed As Boolean = True) As String
        Dim d = DateDiff(interval_, date_, Date.Parse(Now))
        If suffixed Then
            Return ToPlural(d, "days") & " ago"
        Else
            Return d
        End If
    End Function
    'Public Enum FullNameMode
    '	LastNameFirst
    '	FirstNameFirst
    'End Enum
    Public Shared Function Acronym(string_ As String, Optional return_upper_case As Boolean = True, Optional separator As String = " ")
        If string_.Length < 1 Or separator.Length < 1 Then Return ""

        Dim s As Array = SplitTextInSplits(string_, separator, SideToReturn.AsArray)
        Dim r As String = ""
        For i = 0 To s.Length - 1
            r &= Mid(s(i), 1, 1)
        Next
        If return_upper_case Then
            Return r.ToUpper
        Else
            Return r
        End If
    End Function
    'Public Shared Function FullName(last_name As String, first_name As String, middle_name As String, Optional title_of_courtesy As String = Nothing, Optional precede_with_title As Boolean = True, Optional which_name_first As FullNameMode = FullNameMode.LastNameFirst, Optional include_middle_name As Boolean = True) As String
    '	Return NameToFull(last_name, first_name, middle_name, title_of_courtesy, precede_with_title, which_name_first, include_middle_name)
    'End Function

    'Public Shared Function NameToFull(last_name As String, first_name As String, middle_name As String, Optional title_of_courtesy As String = Nothing, Optional precede_with_title As Boolean = True, Optional which_name_first As FullNameMode = FullNameMode.LastNameFirst, Optional include_middle_name As Boolean = True) As String
    '	Dim r As String = ""
    '	If which_name_first = FullNameMode.FirstNameFirst Then
    '		If precede_with_title Then r &= title_of_courtesy & " "
    '		r &= first_name
    '		If middle_name.Length > 0 And include_middle_name = True Then r &= " " & middle_name
    '		r &= " " & last_name
    '	ElseIf which_name_first = FullNameMode.LastNameFirst Then
    '		If precede_with_title Then r &= title_of_courtesy & " "
    '		r &= last_name
    '		r &= " " & first_name
    '		If middle_name.Length > 0 And include_middle_name = True Then r &= " " & middle_name
    '	End If
    '	Return r
    'End Function


    Public Shared Function ToPlural(count_ As Long, str_to_change As SingularWord, Optional rest_of_full_string As String = "", Optional prefixed As Boolean = True, Optional textCase As TextCase = TextCase.LowerCase, Optional prepend_with_approximately_if_count_is_of_type_double As Boolean = False) As String

        If prepend_with_approximately_if_count_is_of_type_double = True And count_ Mod 2 <> 0 Then
            Return "approximately " & ToPlural(count_, str_to_change.ToString, rest_of_full_string, prefixed, textCase)
        Else
            Return ToPlural(count_, str_to_change.ToString, rest_of_full_string, prefixed, textCase)
        End If
    End Function

    Private Shared Function ToPlural(count_ As Long, str_to_change As String, Optional rest_of_full_string As String = "", Optional prefixed As Boolean = True, Optional textCase As TextCase = TextCase.LowerCase) As String
        Dim c = 0
        Try
            If count_.ToString.Length > 0 Then c = count_
        Catch
        End Try

        Dim val_ As String = ""

        If Val(c) = 1 Then
            Select Case str_to_change.ToString.ToLower
                Case "message"
                    If prefixed Then
                        val_ = "a new message"
                    Else
                        val_ = "message"
                    End If
                Case "messages"
                    If prefixed Then
                        val_ = "a new message"
                    Else
                        val_ = "message"
                    End If
                Case "mark"
                    If prefixed Then
                        val_ = "1 mark "
                    Else
                        val_ = "mark"
                    End If
                Case "marks"
                    If prefixed Then
                        val_ = "1 mark "
                    Else
                        val_ = "mark"
                    End If
                Case "minute"
                    If prefixed Then
                        val_ = "1 minute "
                    Else
                        val_ = "minute"
                    End If
                Case "minutes"
                    If prefixed Then
                        val_ = "1 minute "
                    Else
                        val_ = "minute"
                    End If

                Case "day"
                    If prefixed Then
                        val_ = "1 day "
                    Else
                        val_ = "day"
                    End If
                Case "days"
                    If prefixed Then
                        val_ = "1 day "
                    Else
                        val_ = "day"
                    End If

                Case "student"
                    If prefixed Then
                        val_ = "1 student "
                    Else
                        val_ = "student"
                    End If
                Case "students"
                    If prefixed Then
                        val_ = "1 student "
                    Else
                        val_ = "student"
                    End If

                Case "month"
                    If prefixed Then
                        val_ = "1 month "
                    Else
                        val_ = "month"
                    End If
                Case "months"
                    If prefixed Then
                        val_ = "1 month "
                    Else
                        val_ = "month"
                    End If

                Case "year"
                    If prefixed Then
                        val_ = "1 year "
                    Else
                        val_ = "year"
                    End If
                Case "years"
                    If prefixed Then
                        val_ = "1 year "
                    Else
                        val_ = "year"
                    End If

                Case "hour"
                    If prefixed Then
                        val_ = "1 hour "
                    Else
                        val_ = "hour"
                    End If
                Case "hours"
                    If prefixed Then
                        val_ = "1 hour "
                    Else
                        val_ = "hour"
                    End If

                Case "second"
                    If prefixed Then
                        val_ = "1 second "
                    Else
                        val_ = "second"
                    End If
                Case "seconds"
                    If prefixed Then
                        val_ = "1 second "
                    Else
                        val_ = "second"
                    End If

                Case "account"
                    If prefixed Then
                        val_ = "1 account "
                    Else
                        val_ = "account"
                    End If
                Case "accounts"
                    If prefixed Then
                        val_ = "1 account "
                    Else
                        val_ = "account"
                    End If

                Case "male"
                    If prefixed Then
                        val_ = "1 male "
                    Else
                        val_ = "male"
                    End If
                Case "males"
                    If prefixed Then
                        val_ = "1 male "
                    Else
                        val_ = "male"
                    End If

                Case "female"
                    If prefixed Then
                        val_ = "1 female "
                    Else
                        val_ = "female"
                    End If
                Case "females"
                    If prefixed Then
                        val_ = "1 female "
                    Else
                        val_ = "female"
                    End If

                Case "user"
                    If prefixed Then
                        val_ = "1 user "
                    Else
                        val_ = "user"
                    End If
                Case "users"
                    If prefixed Then
                        val_ = "1 user "
                    Else
                        val_ = "user"
                    End If

                Case "file"
                    If prefixed Then
                        val_ = "1 file "
                    Else
                        val_ = "file"
                    End If
                Case "files"
                    If prefixed Then
                        val_ = "1 file "
                    Else
                        val_ = "file"
                    End If

                Case "record"
                    If prefixed Then
                        val_ = "1 record "
                    Else
                        val_ = "record"
                    End If
                Case "records"
                    If prefixed Then
                        val_ = "1 record "
                    Else
                        val_ = "record"
                    End If

                Case "question"
                    If prefixed Then
                        val_ = "1 question "
                    Else
                        val_ = "question"
                    End If
                Case "questions"
                    If prefixed Then
                        val_ = "1 question "
                    Else
                        val_ = "question"
                    End If

                Case "task"
                    If prefixed Then
                        val_ = "1 task "
                    Else
                        val_ = "task"
                    End If
                Case "tasks"
                    If prefixed Then
                        val_ = "1 task "
                    Else
                        val_ = "task"
                    End If

                Case "item"
                    If prefixed Then
                        val_ = "1 item "
                    Else
                        val_ = "item"
                    End If
                Case "items"
                    If prefixed Then
                        val_ = "1 item "
                    Else
                        val_ = "item"
                    End If

                Case "unit"
                    If prefixed Then
                        val_ = "1 unit "
                    Else
                        val_ = "unit"
                    End If
                Case "units"
                    If prefixed Then
                        val_ = "1 unit "
                    Else
                        val_ = "unit"
                    End If

                Case "room"
                    If prefixed Then
                        val_ = "1 room "
                    Else
                        val_ = "room"
                    End If
                Case "rooms"
                    If prefixed Then
                        val_ = "1 room "
                    Else
                        val_ = "room"
                    End If

                Case "client"
                    If prefixed Then
                        val_ = "1 client "
                    Else
                        val_ = "client"
                    End If
                Case "clients"
                    If prefixed Then
                        val_ = "1 client "
                    Else
                        val_ = "client"
                    End If

                Case "match"
                    If prefixed Then
                        val_ = "1 match "
                    Else
                        val_ = "match"
                    End If
                Case "matches"
                    If prefixed Then
                        val_ = "1 match "
                    Else
                        val_ = "match"
                    End If

                Case "invoice"
                    If prefixed Then
                        val_ = "1 invoice "
                    Else
                        val_ = "invoice"
                    End If
                Case "invoices"
                    If prefixed Then
                        val_ = "1 invoice "
                    Else
                        val_ = "invoice"
                    End If

                Case "sale"
                    If prefixed Then
                        val_ = "1 sale "
                    Else
                        val_ = "sale"
                    End If
                Case "sales"
                    If prefixed Then
                        val_ = "1 sale "
                    Else
                        val_ = "sale"
                    End If

                Case "receipt"
                    If prefixed Then
                        val_ = "1 receipt "
                    Else
                        val_ = "receipt"
                    End If
                Case "receipts"
                    If prefixed Then
                        val_ = "1 receipt "
                    Else
                        val_ = "receipt"
                    End If

                Case "product"
                    If prefixed Then
                        val_ = "1 product "
                    Else
                        val_ = "product"
                    End If
                Case "products"
                    If prefixed Then
                        val_ = "1 product "
                    Else
                        val_ = "product"
                    End If

                Case "application"
                    If prefixed Then
                        val_ = "1 application "
                    Else
                        val_ = "application"
                    End If
                Case "applications"
                    If prefixed Then
                        val_ = "1 application "
                    Else
                        val_ = "application"
                    End If

            End Select
        Else
            Select Case str_to_change.ToString.ToLower
                Case "message"
                    If prefixed Then
                        val_ = c & " new messages"
                    Else
                        val_ = "messages"
                    End If
                Case "messages"
                    If prefixed Then
                        val_ = c & " new messages"
                    Else
                        val_ = "messages"
                    End If
                Case "mark"
                    If prefixed Then
                        val_ = c & " marks"
                    Else
                        val_ = "marks"
                    End If
                Case "marks"
                    If prefixed Then
                        val_ = c & " marks"
                    Else
                        val_ = "marks"
                    End If
                Case "minute"
                    If prefixed Then
                        val_ = c & " minutes"
                    Else
                        val_ = "minutes"
                    End If
                Case "minutes"
                    If prefixed Then
                        val_ = c & " minutes"
                    Else
                        val_ = "minutes"
                    End If

                Case "day"
                    If prefixed Then
                        val_ = c & " days"
                    Else
                        val_ = "days"
                    End If
                Case "days"
                    If prefixed Then
                        val_ = c & " days"
                    Else
                        val_ = "days"
                    End If

                Case "student"
                    If prefixed Then
                        val_ = c & " students"
                    Else
                        val_ = "students"
                    End If
                Case "students"
                    If prefixed Then
                        val_ = c & " students"
                    Else
                        val_ = "students"
                    End If

                Case "month"
                    If prefixed Then
                        val_ = c & " months"
                    Else
                        val_ = "months"
                    End If
                Case "months"
                    If prefixed Then
                        val_ = c & " months"
                    Else
                        val_ = "months"
                    End If
                Case "year"
                    If prefixed Then
                        val_ = c & " years"
                    Else
                        val_ = "years"
                    End If
                Case "years"
                    If prefixed Then
                        val_ = c & " years"
                    Else
                        val_ = "years"
                    End If
                Case "hour"
                    If prefixed Then
                        val_ = c & " hours"
                    Else
                        val_ = "hours"
                    End If
                Case "hours"
                    If prefixed Then
                        val_ = c & " hours"
                    Else
                        val_ = "hours"
                    End If
                Case "second"
                    If prefixed Then
                        val_ = c & " seconds"
                    Else
                        val_ = "seconds"
                    End If
                Case "seconds"
                    If prefixed Then
                        val_ = c & " seconds"
                    Else
                        val_ = "seconds"
                    End If
                Case "account"
                    If prefixed Then
                        val_ = c & " accounts"
                    Else
                        val_ = "accounts"
                    End If
                Case "accounts"
                    If prefixed Then
                        val_ = c & " accounts"
                    Else
                        val_ = "accounts"
                    End If

                Case "male"
                    If prefixed Then
                        val_ = c & " males"
                    Else
                        val_ = "males"
                    End If
                Case "males"
                    If prefixed Then
                        val_ = c & " males"
                    Else
                        val_ = "males"
                    End If
                Case "female"
                    If prefixed Then
                        val_ = c & " females"
                    Else
                        val_ = "females"
                    End If
                Case "females"
                    If prefixed Then
                        val_ = c & " females"
                    Else
                        val_ = "females"
                    End If
                Case "user"
                    If prefixed Then
                        val_ = c & " users"
                    Else
                        val_ = "users"
                    End If
                Case "users"
                    If prefixed Then
                        val_ = c & " users"
                    Else
                        val_ = "users"
                    End If
                Case "file"
                    If prefixed Then
                        val_ = c & " files"
                    Else
                        val_ = "files"
                    End If
                Case "files"
                    If prefixed Then
                        val_ = c & " files"
                    Else
                        val_ = "files"
                    End If
                Case "record"
                    If prefixed Then
                        val_ = c & " records"
                    Else
                        val_ = "records"
                    End If
                Case "records"
                    If prefixed Then
                        val_ = c & " records"
                    Else
                        val_ = "records"
                    End If

                Case "question"
                    If prefixed Then
                        val_ = c & " questions"
                    Else
                        val_ = "questions"
                    End If
                Case "questions"
                    If prefixed Then
                        val_ = c & " questions"
                    Else
                        val_ = "questions"
                    End If

                Case "task"
                    If prefixed Then
                        val_ = c & " tasks"
                    Else
                        val_ = "tasks"
                    End If
                Case "tasks"
                    If prefixed Then
                        val_ = c & " tasks"
                    Else
                        val_ = "tasks"
                    End If

                Case "item"
                    If prefixed Then
                        val_ = c & " items"
                    Else
                        val_ = "items"
                    End If
                Case "items"
                    If prefixed Then
                        val_ = c & " items"
                    Else
                        val_ = "items"
                    End If

                Case "unit"
                    If prefixed Then
                        val_ = c & " units"
                    Else
                        val_ = "units"
                    End If
                Case "units"
                    If prefixed Then
                        val_ = c & " units"
                    Else
                        val_ = "units"
                    End If

                Case "room"
                    If prefixed Then
                        val_ = c & " rooms"
                    Else
                        val_ = "rooms"
                    End If
                Case "rooms"
                    If prefixed Then
                        val_ = c & " rooms"
                    Else
                        val_ = "rooms"
                    End If

                Case "client"
                    If prefixed Then
                        val_ = c & " clients"
                    Else
                        val_ = "clients"
                    End If
                Case "clients"
                    If prefixed Then
                        val_ = c & " clients"
                    Else
                        val_ = "clients"
                    End If

                Case "match"
                    If prefixed Then
                        val_ = c & " matches"
                    Else
                        val_ = "matches"
                    End If
                Case "matches"
                    If prefixed Then
                        val_ = c & " matches"
                    Else
                        val_ = "matches"
                    End If

                Case "invoice"
                    If prefixed Then
                        val_ = c & " invoices"
                    Else
                        val_ = "invoices"
                    End If
                Case "invoices"
                    If prefixed Then
                        val_ = c & " invoices"
                    Else
                        val_ = "invoices"
                    End If

                Case "sale"
                    If prefixed Then
                        val_ = c & " sales"
                    Else
                        val_ = "sales"
                    End If
                Case "sales"
                    If prefixed Then
                        val_ = c & " sales"
                    Else
                        val_ = "sales"
                    End If

                Case "receipt"
                    If prefixed Then
                        val_ = c & " receipts"
                    Else
                        val_ = "receipts"
                    End If
                Case "receipts"
                    If prefixed Then
                        val_ = c & " receipts"
                    Else
                        val_ = "receipts"
                    End If

                Case "product"
                    If prefixed Then
                        val_ = c & " products"
                    Else
                        val_ = "products"
                    End If
                Case "products"
                    If prefixed Then
                        val_ = c & " products"
                    Else
                        val_ = "products"
                    End If

                Case "application"
                    If prefixed Then
                        val_ = c & " applications"
                    Else
                        val_ = "applications"
                    End If
                Case "applications"
                    If prefixed Then
                        val_ = c & " applications"
                    Else
                        val_ = "applications"
                    End If

            End Select
        End If

        If rest_of_full_string.Length > 0 Then
            val_ &= " " & rest_of_full_string
        End If

        Return TransformText(val_, textCase)
    End Function

    ''' <summary>
    ''' Determines if val_ is within 0 and 100.
    ''' </summary>
    ''' <param name="val_"></param>
    ''' <returns></returns>
    Public Shared Function InPercent(val_ As String) As Boolean
        val_ = val_.Trim
        If Val(val_) < 0 Or Val(val_) > 100 Then ' Or Val(val_) = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function IsEven(val As Long) As Boolean
        IsEven = (val Mod 2 = 0)
    End Function
    Public Shared Function ToBString(val) As String
        Select Case val.ToString
            Case True
                Return "Yes"
            Case False
                Return "No"
        End Select
    End Function
    Public Shared Function GetMonth(number) As String
        Return MonthString(number)
    End Function

    Public Shared Function MonthString(number) As String
        Dim result As String = ""
        Select Case number
            Case 1
                result = "January"
            Case 2
                result = "February"
            Case 3
                result = "March"
            Case 4
                result = "April"
            Case 5
                result = "May"
            Case 6
                result = "June"
            Case 7
                result = "July"
            Case 8
                result = "August"
            Case 9
                result = "September"
            Case 10
                result = "October"
            Case 11
                result = "November"
            Case 12
                result = "December"
        End Select
        Return result
    End Function

#End Region

#Region "Machine"

    ''' <summary>
    ''' Checks if PC is plugged in.
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function MachineIsOnAC() As Boolean
        Dim myManagedPower As New ManagedPower()
        Return myManagedPower.ToString().ToLower = "ac"
    End Function
    ''' <summary>
    ''' Checks if PC is not plugged in (i.e. running on battery).
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function MachineIsOnBattery() As Boolean
        Dim myManagedPower As New ManagedPower()
        Return myManagedPower.ToString().ToLower <> "ac"
    End Function
    Public Shared Function ReplaceText(body_of_text As String, word_to_remove As String, Optional what_to_replace_with As String = "")
        Return body_of_text.Replace(word_to_remove, what_to_replace_with)
    End Function
    Public Shared Function ReplaceText(body_of_text As String, list_of_words_to_remove As Array, Optional what_to_replace_with As String = "")
        Dim current_result As String
        Dim result As String = body_of_text
        For i = 0 To list_of_words_to_remove.Length - 1
            current_result = result.Replace(list_of_words_to_remove(i), what_to_replace_with)
            result = current_result
        Next
        Return result
    End Function

    Public Shared Function ReplaceText(body_of_text As String, list_of_words_to_remove As List(Of String), Optional what_to_replace_with As String = "")
        Return ReplaceText(body_of_text, list_of_words_to_remove.ToArray, what_to_replace_with)
    End Function

    '''' <summary>
    '''' bugged. do not use.
    '''' </summary>
    '''' <param name="text_to_place_on_clipboard"></param>
    'Private Shared Sub WriteText(text_to_place_on_clipboard As String)
    '	Try
    '		Clipboard.SetText(text_to_place_on_clipboard)
    '	Catch ex As Exception
    '	End Try
    'End Sub


    Public Shared Shadows Function Write(text As String, Optional write_as As WriteAs = WriteAs.General, Optional formatForSQL As Boolean = False) As String

        Dim result As String = ""
        If text.Length > 0 Then
            Select Case write_as
                Case WriteAs.AsIs
                    result = text
                Case WriteAs.General
                    result = text.Replace(vbCrLf, "<br />")
            End Select
        End If
        If formatForSQL Then
            result = result.Replace("'", "''")
        End If
        Return result
    End Function




    ''' <summary>
    ''' Creates, adds to, or overwrites the content of file (format is text).
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="txt_">Intended string content of the file.</param>
    ''' <param name="append_">Should it add to the content of the file (if it has) or overwrite everything?</param>
    ''' <param name="dont_trim">Should trailing spaces be ignored?</param>
    Public Shared Function WriteText(file_ As String, txt_ As String, Optional append_ As Boolean = False, Optional dont_trim As Boolean = True)
        'If file_.Length < 1 Then Return

        'Dim t As String = txt_
        'If dont_trim = False Then t = txt_.Trim
        'Try
        '	My.Computer.FileSystem.WriteAllText(file_, t, append_)
        'Catch ex As Exception
        'End Try



        Dim [FileVar] As String = file_
        If append_ = True Then
            FileOpen(1, [FileVar], OpenMode.Append)
        Else
            FileOpen(1, [FileVar], OpenMode.Output)
        End If
        If dont_trim Then
            PrintLine(1, txt_)
        Else
            PrintLine(1, txt_.Trim)
        End If
        FileClose(1)

    End Function

    ''' <summary>
    ''' Retrieves the content of a file (format is text).
    ''' </summary>
    ''' <param name="file_">The path to the file.</param>
    ''' <param name="dont_trim">Should trailing spaces be ignored?</param>
    ''' <returns>The (text) content of the file.</returns>
    Public Shared Function ReadText(file_ As String, Optional dont_trim As Boolean = False) As String
        'If file_ Is Nothing Then Exit Function

        'If CType(file_, String).Length < 1 Then Exit Function

        If file_.Length < 1 Then Return ""
        If Path.GetExtension(file_).Contains("doc") Then GoTo 2

        Dim r As String = My.Computer.FileSystem.ReadAllText(file_).Trim
        If dont_trim = True Then
            r = My.Computer.FileSystem.ReadAllText(file_)
        End If
        Return r




        'Dim docName As String = Path.GetFileName(file_)
        'Dim docPath As String = Path.GetDirectoryName(file_)
        'Dim stream As New FileStream(file_, FileMode.Open)
        'Dim reader As New StreamReader(stream)
        'Try
        '	If dont_trim = True Then
        '		Return reader.ReadToEnd()
        '	Else
        '		Return reader.ReadToEnd().Trim
        '	End If
        'Catch
        'Finally
        '	reader.Dispose()
        '	stream.Dispose()
        'End Try
        'Return True


2:

    End Function



    '''' <summary>
    '''' bugged. do not use.
    '''' </summary>
    '''' <param name="control_"></param>
    '''' <returns></returns>
    'Private Shared Function ReadText(Optional control_ As Object = Nothing) As String
    '	''		If control_ Is Nothing Then Exit Function

    '	Dim s As String = ""
    '	Try
    '		s = Clipboard.GetText
    '	Catch ex As Exception
    '	End Try
    '	Try
    '		control_.text = s
    '	Catch ex As Exception
    '	End Try
    '	Return s
    'End Function


    ''' <summary>
    ''' Constructs a preview HTML file
    ''' </summary>
    ''' <param name="content"></param>
    ''' <param name="title"></param>
    ''' <returns></returns>
    Public Shared Function WriteHTML(content As String, Optional title As String = "Preview") As String
        Return PreviewHTMLContent(content, title)
    End Function

    ''' <summary>
    ''' Constructs and creates a preview HTML file, and optionally accompanies it with static resources (Shoppy).
    ''' </summary>
    ''' <param name="content"></param>
    ''' <param name="directory"></param>
    ''' <param name="include_static_resources"></param>
    ''' <param name="css_folder"></param>
    ''' <param name="fonts_folder"></param>
    ''' <param name="js_folder"></param>
    ''' <param name="title"></param>
    ''' <param name="filename_without_extension"></param>
    Public Shared Sub WriteHTML(content As String, directory As String, Optional include_static_resources As Boolean = True, Optional css_folder As String = Nothing, Optional fonts_folder As String = Nothing, Optional js_folder As String = Nothing, Optional title As String = "Preview", Optional filename_without_extension As String = "index")
        If directory.Length < 1 Then Return

        Dim result As String = PreviewHTMLContent(content, title)

        Try
            MkDir(directory)
        Catch ex As Exception

        End Try

        Dim file As String = directory & "\" & filename_without_extension & ".html"
        Try
            WriteText(file, result)
        Catch ex As Exception

        End Try

        If include_static_resources = True And
                css_folder.Length > 0 And
                js_folder.Length > 0 And
                fonts_folder.Length > 0 Then
            Try
                My.Computer.FileSystem.CopyDirectory(css_folder, directory & "\css", FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
                My.Computer.FileSystem.CopyDirectory(js_folder, directory & "\js", FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
                My.Computer.FileSystem.CopyDirectory(fonts_folder, directory & "\fonts", FileIO.UIOption.AllDialogs, FileIO.UICancelOption.DoNothing)
            Catch ex As Exception

            End Try
        End If

    End Sub


    Public Shared Function PreviewHTMLContent(content As String, Optional title As String = "Preview") As String

        If content.Length > 0 Then
            content = content.Replace(vbCrLf, "<br />")
        End If

        Dim result As String = "<!DOCTYPE html>
<html lang=""en"">

<head>
    <meta charset=""UTF-8"">
    <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>" & title & "</title>

    <script type=""application/x-javascript"">
		addEventListener(""load"", function() { setTimeout(hideURLbar, 0); }, false); function hideURLbar(){ window.scrollTo(0,1); }
        </script>
    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <link href=""css/bootstrap.css"" rel=""stylesheet"" type=""text/css"" media=""all"" />
    <!-- Custom Theme files -->
    <link href=""css/style.css"" rel=""stylesheet"" type=""text/css"" media=""all"" />
    <!--js-->
    <script src=""js/jquery-2.1.1.min.js""></script>
    <!--icons-css-->
    <!--<link href=""css/font-awesome.css"" rel=""stylesheet"" />-->


    <link rel=""stylesheet"" type=""text/css"" href=""fonts/font-awesome-4.7.0/css/font-awesome.min.css"" />

    <!--Google Fonts-->
    <link href='//fonts.googleapis.com/css?family=Carrois+Gothic' rel='stylesheet' type='text/css' />
    <link href='//fonts.googleapis.com/css?family=Work+Sans:400,500,600' rel='stylesheet' type='text/css' />
    <!--//skycons-icons-->
    <!--button css-->
    <link href=""css/demo-page.css"" rel=""stylesheet"" media=""all"" />
    <link href=""css/hover.css"" rel=""stylesheet"" media=""all"" />
    <!--static chart-->
    <script src=""js/Chart.min.js""></script>
    <!--//charts-->
    <!-- geo chart
	<script src=""//cdn.jsdelivr.net/modernizr/2.8.3/modernizr.min.js"" type=""text/javascript""></script>
	rem removed-->
    <script src=""js/modernizr.min.js""></script>
    <!-- rem added  -->

    <script src=""lib/modernizr/modernizr-custom.js""></script>
    <!-- rem added  -->
    
    
    
    
    
    <script>
        window.modernizr || document.write('<script src=""lib/modernizr/modernizr-custom.js""><\/script>')
    </script>

    <!--<script src=""lib/html5shiv/html5shiv.js""></script>-->
    <!-- Chartinator  -->
    <script src=""js/chartinator.js""></script>
    <!--geo chart-->

    <!--skycons-icons-->
    <script src=""js/skycons.js""></script>
    <!--//skycons-icons-->
    <!--mapcss-->
    <link rel=""stylesheet"" type=""text/css"" href=""css/examples.css"" />
    <!--js-->
    <script type=""text/javascript"" src=""//maps.google.com/maps/api/js?sensor=true""></script>
    <script type=""text/javascript"" src=""js/gmaps.js""></script>
    <!--map-->
    <!--pop up strat here-->
    <script src=""js/jquery.magnific-popup.js"" type=""text/javascript""></script>
    <script>
        $(document).ready(function () {
            $('.popup-with-zoom-anim').magnificPopup({
                type: 'inline',
                fixedContentPos: false,
                fixedBgPos: true,
                overflowY: 'auto',
                closeBtnInside: true,
                preloader: false,
                midClick: true,
                removalDelay: 300,
                mainClass: 'my-mfp-zoom-in'
            });

        });
    </script>

    <link rel=""stylesheet"" href=""css/inovation/_style.css"">
</head>

<body>

    <div class=""navbar navbar-inverse navbar-fixed-top"">
        <div class=""container"">
            <div class=""navbar-header"">
                <button type=""button"" class=""navbar-toggle"" data-toggle=""collapse"" data-target="".navbar-collapse""
                    title=""more options"">
                    <span class=""icon-bar""></span>
                    <span class=""icon-bar""></span>
                    <span class=""icon-bar""></span>
                </button>
                <a class=""navbar-brand"" runat=""server"" href=""#"">
                    Preview
                </a>
            </div>
        </div>
    </div>

    <div class=""container"" style=""padding-top: 80px"">
        <div class=""row"">
            <div class=""col col-sm-12 col-md-2"">
            </div>
            <div class=""col col-sm-12 col-md-8"">" & content & "</div>
            <div class=""col col-sm-12 col-md-2"">
            </div>

        </div>
    </div>



    <!--scrolling js-->
    <script src=""js/jquery.nicescroll.js""></script>
    <script src=""js/scripts.js""></script>
    <!--//scrolling js-->
    <script src=""js/bootstrap.js""></script>
    <script type=""text/javascript"">
        function ScrollToDIV(div_) {
            div_.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    </script>
</body>

</html>"

        Return result
    End Function

    Public Shared Function ToPath(str_ As String) As String
        Try
            If str_.Substring(str_.Length - 1, 1) <> "\" Then
                Return str_ & "\"
            Else
                Return str_
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Shared Function IsImage(filename_ As String) As Boolean
        Select Case System.IO.Path.GetExtension(filename_).ToLower
            Case ".jpg"
                Return True
            Case ".jpeg"
                Return True
            Case ".png"
                Return True
            Case ".gif"
                Return True
            Case ".bmp"
                Return True
            Case ".tif"
                Return True
            Case ".ico"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Shared Sub Delete(file_ As String, Optional recycle_ As Boolean = True, Optional showUI_ As Boolean = False)
        Try
            My.Computer.FileSystem.DeleteFile(file_, showUI_, recycle_)
        Catch ex As Exception
        End Try
    End Sub
    Public Shared Sub Delete(files_ As Array, Optional recycle_ As Boolean = True, Optional showUI_ As Boolean = False)
        If files_.Length > 0 Then
            For i = 0 To files_.Length - 1
                Delete(files_(i), recycle_, showUI_)
            Next
        End If
    End Sub



    Private Shared Function array_to_string(array As Array) As String
        Dim s = ""
        Try
            For i = 0 To array.Length - 1
                s &= array(i) & " "
            Next
        Catch
        End Try
        Return s.Trim
    End Function

#End Region


    '#Region "Scanner"
    '    Public Shared Function ScanText(source_file As String, Optional target_filename_full_path As String = Nothing)

    '        Try
    '            Kill(target_filename_full_path)
    '        Catch
    '        End Try

    '        Try
    '            Dim md As MODI.Document = New MODI.Document()
    '            md.Create(Convert.ToString(source_file))
    '            md.OCR(MODI.MiLANGUAGES.miLANG_ENGLISH, True, True)
    '            Dim image As MODI.Image = CType(md.Images(0), MODI.Image)
    '            Try
    '                Dim createFile As FileStream = New FileStream(target_filename_full_path, FileMode.CreateNew)
    '                Dim writeFile As StreamWriter = New StreamWriter(createFile)
    '                writeFile.Write(image.Layout.Text)
    '                writeFile.Close()
    '            Catch
    '            End Try
    '            Return image.Layout.Text
    '        Catch exc As Exception
    '        End Try
    '    End Function

    '#End Region

#Region "SJ"

    ''' <summary>
    ''' Ideal for GeneralModule.FormatWindow.StartAppsWithArguments
    ''' </summary>
    ''' <param name="keys_"></param>
    ''' <param name="values_"></param>
    ''' <returns></returns>
    Public Shared Function Si_ListsToKeyPairValuePair(keys_ As List(Of String), values_ As List(Of String)) As String
        '		Dim file_2 As String = "'1':'notepad', '2':'C:\Users\Administrator\Desktop\Hub\2D\StrictD\1.txt', '3':'C:\Program Files (x86)\Windows Media Player\wmplayer.exe', '4':'C:\Users\Pediforte\Music\Music\Playlists\68.wpl'"
        Dim s As String = ""
        Dim counter As Integer = 1
        With keys_
            For i = 0 To .Count - 1
                s &= "'" & counter & "':'" & keys_(i) & "', '" & counter + 1 & "':'" & values_(i) & "'"
                counter += 2
                If i < .Count - 1 Then s &= ", "
            Next
        End With
        Return s
    End Function

    ''' <summary>
    ''' Turns Array, Collection, List (Of String, Object, Integer, Double) to SEMI-JSON or JSON if as_json.
    ''' </summary>
    ''' <param name="keys_values__"></param>
    ''' <param name="as_json"></param>
    ''' <returns></returns>
    Public Shared Function Si_ListToJSON(keys_values__ As Object, Optional as_json As Boolean = False) As String
        Dim keys_values
        Dim key_values_array As Array
        If TypeOf keys_values__ Is Collection Then
            keys_values = CType(keys_values__, Collection)
        ElseIf TypeOf keys_values__ Is List(Of String) Then
            keys_values = CType(keys_values__, List(Of String))
        ElseIf TypeOf keys_values__ Is List(Of Object) Then
            keys_values = CType(keys_values__, List(Of Object))
        ElseIf TypeOf keys_values__ Is List(Of Integer) Then
            keys_values = CType(keys_values__, List(Of Integer))
        ElseIf TypeOf keys_values__ Is List(Of Double) Then
            keys_values = CType(keys_values__, List(Of Double))
        ElseIf TypeOf keys_values__ Is Array Then
            key_values_array = CType(keys_values__, Array)
        End If

        Dim s As String = ""

        If TypeOf keys_values__ Is Array Then
            With key_values_array
                If .Length < 1 Then Return ""
                For i As Integer = 0 To .Length - 2 Step 2
                    s &= "'" & key_values_array(i) & "':'" & key_values_array(i + 1) & "'"
                    If i <> .Length - 2 Then s &= ", "
                Next
            End With
        ElseIf TypeOf keys_values__ Is Collection Then
            If keys_values.Count < 1 Then Return ""
            With keys_values
                For i As Integer = 1 To .Count - 1 Step 2
                    s &= "'" & keys_values(i) & "':'" & keys_values(i + 1) & "'"
                    If i <> .Count - 1 Then s &= ", "
                Next
            End With
        Else
            If keys_values.Count < 1 Then Return ""
            With keys_values
                For i As Integer = 0 To .Count - 2 Step 2
                    s &= "'" & keys_values(i) & "':'" & keys_values(i + 1) & "'"
                    If i <> .Count - 2 Then s &= ", "
                Next
            End With
        End If
        If as_json Then
            Return Si_ToJSON(s)
        Else
            Return s
        End If
    End Function

    Public Shared Function jsonListToJSON(dict As Dictionary(Of String, String), Optional keys_are_numbers As Boolean = False, Optional as_json As Boolean = False) As String
        Return Si_ListToJSON(dict, keys_are_numbers, as_json)
    End Function

    Public Shared Function Si_ListToJSON(dict As Dictionary(Of String, String), Optional keys_are_numbers As Boolean = False, Optional as_json As Boolean = False) As String

        Dim s As String = ""

        If dict.Count < 1 Then Return s

        If keys_are_numbers Then
            For i = 0 To dict.Count - 1
                s &= "'" & i + 1 & "':'" & dict.Values(i) & "'"
                If i < dict.Count - 1 Then
                    s &= ", "
                End If
            Next
        Else
            For i = 0 To dict.Count - 1
                s &= "'" & dict.Keys(i).ToString & "':'" & dict.Values(i) & "'"
                If i < dict.Count - 1 Then
                    s &= ", "
                End If
            Next
        End If

        If as_json Then
            Return Si_ToJSON(s)
        Else
            Return s
        End If
    End Function

    ''' <summary>
    ''' Turns single quoted, comma-separated values (w/wo preceding { or trailing }) to list of values. Just bind the list it returns to the ListBox or ComboBox directly, or loop through to populate the ListBox, ComboBox or TextBox (using D.Si_BindListToControl()).
    ''' </summary>
    ''' <param name="val_">Database field's string</param>
    ''' <returns>List you can bind to ListBox or ComboBox</returns>
    Public Shared Function Si_StringToList(val_ As String) As List(Of String)
        Dim l As New List(Of String)
        Try
            l.Clear()
        Catch ex As Exception
        End Try
        If val_.Trim.Length < 1 Then
            Return l
            Exit Function
        End If

        Dim str_ As String = ""
        If Mid(val_.Trim, 1, 1) <> "{" Then
            str_ = "{" & val_.Trim & "}"
        Else
            str_ = val_.Trim
        End If

        str_ = str_.Replace("\", "\\")

        Dim w As JObject = JObject.Parse(str_)

        Dim s() = w.PropertyValues.ToArray
        For i As Integer = 0 To s.Count - 1
            l.Add(s(i).ToString)
        Next
        Return l
    End Function

    ''' <summary>
    ''' Checks if string contains property.
    ''' </summary>
    ''' <param name="sj_string">JSON string</param>
    ''' <param name="property_">Property name</param>
    ''' <returns>True if property_ with value_ is found, False otherwise</returns>
    Public Shared Function Si_DoesPropertyExist(sj_string As String, property_ As String) As Boolean
        Dim str_ As String
        Dim val_ As String = sj_string
        If Mid(val_.Trim, 1, 1) <> "{" Then
            str_ = "{" & val_.Trim & "}"
        Else
            str_ = val_.Trim
        End If
        Dim w As JObject = JObject.Parse(str_)
        Dim value_ As String = ""
        Return w.TryGetValue(property_, value_)
    End Function

    ''' <summary>
    ''' Gets the value of a property
    ''' </summary>
    ''' <param name="sj_string">JSON string</param>
    ''' <param name="property_">The value of this property will be returned</param>
    ''' <returns>String</returns>
    Public Shared Function Si_GetValue(sj_string As String, property_ As String) As String
        Dim str_ As String = ""
        Dim val_ As String = sj_string
        If Mid(val_.Trim, 1, 1) <> "{" Then
            str_ = "{" & val_.Trim & "}"
        Else
            str_ = val_.Trim
        End If
        Dim w As JObject = JObject.Parse(str_)
        Return w.GetValue(property_)
    End Function

    Public Shared Function Si_HowManyProperties(sj_string As String) As Long
        Dim str_ As String = ""
        Dim val_ As String = sj_string
        If Mid(val_.Trim, 1, 1) <> "{" Then
            str_ = "{" & val_.Trim & "}"
        Else
            str_ = val_.Trim
        End If
        Dim w As JObject = JObject.Parse(str_)
        Return w.Properties.Count
    End Function

    ''' <summary>
    ''' Turns string to SI_JSON.
    ''' </summary>
    ''' <param name="val_">String to turn to JSON</param>
    ''' <returns>JSON</returns>
    Public Shared Function Si_ToJSON(val_) As String
        Dim str_ As String = ""
        If Mid(val_.Trim, 1, 1) <> "{" Then
            str_ = "{" & val_.Trim & "}"
        Else
            str_ = val_.Trim
        End If
        Return str_.Replace("\", "\\")
    End Function

    ''' <summary>
    ''' Gets value from property.
    ''' </summary>
    ''' <param name="val_">Value sought.</param>
    ''' <param name="property_">Property to search for.</param>
    ''' <returns></returns>
    Private Shared Function Si_GetValueJSON(val_ As String, property_ As String) As String
        Dim x As JObject = JObject.Parse(Si_ToJSON(val_))
        Return x.GetValue(property_).ToString
    End Function

    Public Shared Function Si_ConcatJSON(previous_json_string As String, new_json_string As String) As String
        Return previous_json_string & ", " & new_json_string
    End Function

    Public Shared Function Si_ToDateJSON(date_)
        Dim d_ As DateTime = DateTime.Parse(date_)
        Return d_.Year & "-" & LeadingZero(d_.Month) & "-" & LeadingZero(d_.Day)
    End Function

    ''' <summary>
    ''' Turns list(s of properties(keys) and values) to JSON. You can then slot it into the database. For the opposite, you can use DatabaseToListJSON to get the values.
    ''' </summary>
    ''' <param name="l_p">List of properties.</param>
    ''' <param name="l_v">List of values corresponding to properties (keys).</param>
    ''' <returns>SJ string.</returns>
    Public Shared Function Si_DictionaryToDatabaseJSON(l_p As List(Of String), l_v As List(Of String)) As String
        Dim val_ As String = ""
        For i As Integer = 0 To l_p.Count - 1
            val_ &= "'" & l_p.Item(i) & "':'" & l_v.Item(i)
            If i < l_p.Count - 1 Then
                val_ &= "', "
            Else
                val_ &= "'"
            End If
        Next
        Return val_
    End Function

    ''' <summary>
    ''' Turns list to JSON. You can then slot it into the database. For the opposite, you can use DatabaseToListJSON to get the values.
    ''' </summary>
    ''' <param name="l"></param>
    ''' <returns></returns>
    Public Shared Function Si_DictionaryToDatabaseJSON(l As List(Of String)) As String
        Dim val_ As String = ""
        Dim ltemp As New List(Of String)
        Try
            ltemp.Clear()
        Catch ex As Exception
        End Try
        ltemp = l
        If ltemp.Count > 0 Then
            With ltemp
                For k As Integer = 0 To .Count - 1 Step 2
                    val_ &= "'" & .Item(k) & "':'" & .Item(k + 1) '& "'"
                    If k <> .Count - 2 Then
                        val_ &= "', "
                    Else
                        val_ &= "'"
                    End If
                Next
            End With
        End If
        Return val_
    End Function

    Public Enum Expecting
        List
        Dictionary
    End Enum

    ''' <summary>
    ''' Turns grid's values to JSON.
    ''' </summary>
    ''' <param name="grid">GridView with 1 column (for values only) or 2 columns (for keys and values).</param>
    ''' <param name="sending">Take values only or both from grid.</param>
    ''' <param name="as_json">As JSON or Semi-JSON?</param>
    ''' <returns>JSON or SJ string.</returns>
    Public Shared Function Si_GridToJSON(grid As System.Web.UI.WebControls.GridView, Optional sending As Expecting = Expecting.Dictionary, Optional as_json As Boolean = False) As String
        Dim l As New List(Of String)

        Dim g As System.Web.UI.WebControls.GridView
        g = grid
        With g
            If .Rows.Count < 1 Then Return ""
            For i As Integer = 0 To .Rows.Count - 1
                If sending = Expecting.List Then
                    l.Add(i + 1)
                    l.Add(.Rows(i).Cells(0).Text)
                ElseIf sending = Expecting.Dictionary Then
                    l.Add(.Rows(i).Cells(0).Text)
                    l.Add(.Rows(i).Cells(1).Text)
                End If
            Next
        End With
        Return Si_ListToJSON(l, as_json)

    End Function

    ''' <summary>
    ''' Turns grid's values to JSON.
    ''' </summary>
    ''' <param name="grid">GridView with 1 column (for values only) or 2 columns (for keys and values).</param>
    ''' <param name="sending">Take values only or both from grid.</param>
    ''' <param name="as_json">As JSON or Semi-JSON?</param>
    ''' <returns>JSON or SJ string.</returns>
    Public Shared Function Si_GridToJSON(grid As System.Windows.Forms.DataGridView, Optional sending As Expecting = Expecting.Dictionary, Optional as_json As Boolean = False) As String
        Dim allow_add As Boolean = False
        Dim allow_delete As Boolean = False

        If grid.AllowUserToAddRows = True Then allow_add = True
        If grid.AllowUserToDeleteRows = True Then allow_delete = True

        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False

        Dim l As New List(Of String)

        Dim g As System.Windows.Forms.DataGridView
        g = grid
        With g
            If .Rows.Count < 1 Then Return ""
            For i As Integer = 0 To .Rows.Count - 1
                If sending = Expecting.List Then
                    l.Add(i + 1)
                    l.Add(.Rows(i).Cells(0).Value)
                ElseIf sending = Expecting.Dictionary Then
                    l.Add(.Rows(i).Cells(0).Value)
                    l.Add(.Rows(i).Cells(1).Value)
                End If
            Next
        End With
        Dim return_ As String = Si_ListToJSON(l, as_json).Replace(", '':''", "")
        grid.AllowUserToAddRows = allow_add
        grid.AllowUserToDeleteRows = allow_delete
        Return return_
    End Function

    Public Shared Function Si_Properties(si_json_string As String) As List(Of String)
        Dim s As String = Si_ToJSON(si_json_string)
        Dim j As JObject = JObject.Parse(s)
        Dim l As New List(Of String)
        For i = 0 To j.Properties.Count - 1
            l.Add(j.Properties(i).Name.ToString)
        Next
        Return l
    End Function

    Public Shared Function Si_Values(si_json_string As String) As List(Of String)
        Dim l As List(Of String) = Si_Properties(si_json_string)
        Dim v As New List(Of String)
        For i = 0 To l.Count - 1
            v.Add(Si_GetValue(Si_ToJSON(si_json_string), l(i)))
        Next
        Return v
    End Function

#End Region


End Class
