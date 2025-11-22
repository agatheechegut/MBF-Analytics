' **************************************************
' MBF Analytics
' Banque de méthodes de Données financières en ligne
' Simple utilisation
' Robuste
' Personnalisable API Key
' ***************************************************


' References :  Outils > References
' 1- Microsoft HTML Object Library
' 2- Microsoft WinHTTP Services, version 5.1

' Classe
' 1- Webtool
' 2- Dictionnary


' Constantes des la classe
Private Const keylen As Integer = 50

' Propriétés de la classe
Private apikey As String    ' code perso pour se connecter à Rapid API
Private h As Object         ' h objet pour les requêtes HTTP
Public status As String     ' état de la connection avec Rapid API :
                            ' ok : tout va bien
                            ' API Key invalide : problème de clé API
                            ' Compte épuisé : plus de requêtes possible
                            ' Paramètres de requête incohérents : Invalid request
                            ' Endpoint Inconnue : Endpoint * does not exist
                            ' Erreur non reconnue
                            

' ******************
' Méthodes Internes
' ******************
' Constructeur de la classe
Private Sub Class_Initialize()
    Set h = New WebTool
    status = "ok"
End Sub

Public Sub initKey(key As String)
    
    apikey = Trim(key)
    
    If Len(apikey) <> keylen Then
        status = "API Key invalide"
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        End
    Else
        status = "ok"
    End If
        
    Set h = New WebTool

End Sub

' ------------------------------------------------------------
' Fonction de recherche ou création d'une feuille Excel nommée
' Figée
' Paramètre
' - sheetname : nom de la feuille Excel recherchée
' Retour : objet feuille Excel
Private Function getOrCreateSheet(ByVal sheetname As String) As Worksheet
    Dim w As Worksheet
    
    Set w = Nothing
    
    For Each w In Worksheets
        If w.Name = sheetname Then Exit For
    Next w
    
    If w Is Nothing Then
        Set w = Worksheets.Add
        w.Name = sheetname
    End If

    Set getOrCreateSheet = w
End Function


' ***********************************************
' ***********************************************
' Collection de Méthodes financières de la classe
' ***********************************************
' ***********************************************

' ****************************************************************
' Batterie de test des erreurs possibles du résultat des requêtes
' ****************************************************************
' Paramètre
' - resultat : le résultat de la requête RapidAPI
' Retour : MsgBox puis End si une erreur est détectée
Private Sub testErreur(Optional ByVal resultat As String = "")
    
    If resultat Like "*" & "not subscribed * API" & "*" Then status = "Souscription à l'API nécessaire"

    If resultat Like "*" & "Endpoint * does not exist" & "*" Then status = "Endpoint inconnue"
    
    If resultat Like "*" & "Invalid request" & "*" Then status = "Paramètres de requête incohérents"

    If resultat Like "*" & "Wrong * parameter" & "*" Then status = "Paramètres de requête incohérents"
        
    If resultat Like "*" & "data"":null" & "*" Then status = "Pas de résultats. Vérifier les paramètres"
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requêtes dépassé"
    
    ' ajouter ses propres détctions d'erreur ici et signaler l'auteur et l'objectif
    ' By XXXXX Objet YYYYY
    ' If resultat Like "*" & "??? * ???" & "*" Then status = "?????"
        

    If status <> "ok" Then
        MsgBox "Erreur : " & status & vbCrLf & vbCrLf & resultat, vbCritical, "MBF Analytics"
        End
    End If

End Sub


' *************************************
' Composantes financières d'une société
' *************************************
' Auteur : Yves Bertini
' Lien page API : https://rapidapi.com/vibapidev/api/bloomberg-real-time/playground/apiendpoint_ac53fec3-fbdc-4a73-ba91-1217728e55c9
' Params avec valeur d'exemple :
'  sheetname = "bloomberg_financials" nom de la feuille crée
'  symbol ou id = "ORA:FP"  Orange sur bourse de Paris France Paris, ticker Bloomberg
'  currencyname = "EUR"     Euros
' Retour :
'  résultats formatés sur feuille excel sheetname
'
' Exemple invocation Bloomberg Financial
'    Call m.bloomberg_financials(sheetname:="Orange Soc", symbol:="ORA:FP", currencyname:="EUR")
'
Public Sub bloomberg_financials(Optional ByVal sheetname = "bloomberg_financials", Optional ByVal symbol = "ORA:FP", Optional ByVal currencyname = "EUR")
    
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête à l'API
    hote = "bloomberg-real-time.p.rapidapi.com"
    epqs = "/stocks/financials?query=SYMBOL&currency=CURRENCYNAME" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    currencyname = WorksheetFunction.EncodeURL(currencyname)
    
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "CURRENCYNAME", currencyname)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = o("data")(3)("securityID")
    w.[B1] = currencyname
    
    
    ' Puis écrire toutes les informationss
    
    ' Income Statement, Balance Sheet, CashFlow
    n1 = 0
    For k = 1 To o("data").Count        ' tous les for avec extrémité Count
        
        Set co = w.[A2].Rows(2 + n1)  ' cellule origine du tableau
        co.Value = o("data")(k)("name"): n = n + 1
        n2 = 0
        For t = 1 To o("data")(k)("timeBasedSheets").Count
            Set ca = co.Rows(2 + n2)   ' cellule origine des data
            
            ca.Value = o("data")(k)("timeBasedSheets")(t)("name")
            p = 1
            For i = 1 To o("data")(k)("timeBasedSheets")(t)("columnHeadings").Count
                ca.Columns(2).Columns(i) = o("data")(k)("timeBasedSheets")(t)("columnHeadings")(i)
            Next i
            
            Set cd = ca.Rows(2)
            For n = 1 To o("data")(k)("timeBasedSheets")(t)("chartData").Count
                cd.Rows(n) = o("data")(k)("timeBasedSheets")(t)("chartData")(n)("name")
                For i = 1 To o("data")(k)("timeBasedSheets")(t)("chartData")(n)("values").Count
                    cd.Columns(2).Rows(n).Columns(i) = o("data")(k)("timeBasedSheets")(t)("chartData")(n)("values")(i)
                    cd.Columns(2).Rows(n).Columns(i).Style = "Currency"
                Next i
                p = p + 1
            Next n
            n2 = n2 + p + 1
        Next t
        n1 = n1 + n2 + 1
    Next k
    
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
End Sub



' ***************************************
' Historique récent des Cotation de titre
' ***************************************
' Auteur : Yves Bertini
' Lien vers la page API : https://rapidapi.com/StoxApi/api/real-time-quotes1/playground/endpoint_c974b05a-9606-45c7-a541-5295f24857c2
' Params avec valeur d'exemple :
'  sheetname = "real_time_quotes1" nom de la feuille crée
'  interval = "4hour" ou 1min 5min 15min 30min 1hour 4hour
'  symbol = "ORA.PA"  Orange sur bourse de Paris France Paris, ticker Yahoo Finance
'  from  = "2025-01-01"     date de départ ISO en texte, FR UTC+1
'  until  = "2025-11-01"     date de fin ISO en texte, FR UTC+1
' Retour :
'  résultats formatés sur feuille excel sheetname
'
' Exemple invocation  Time Quotes1
' Call m.real_time_quotes1(sheetname:="Orange Sto", interval:="4hour", symbol:="ORA.PA", fromdt:="2025-06-01", untildt:="2025-11-01")
'
Public Sub real_time_quotes1(Optional ByVal sheetname = "real_time_quotes1", Optional ByVal interval = "4hour", Optional ByVal symbol = "ORA.PA", Optional ByVal fromdt = "2025-06-01", Optional ByVal untildt = "2025-11-01")
        
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "real-time-quotes1.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/api/v2/historical/stock?interval=INTERVAL&symbol=SYMBOL&from=FROMDT&until=UNTILDT"
    
    
    ' Intégration des PARAMS dans la Query String
    ' Replace sensible MAJ MIN
    epqs = Replace(epqs, "INTERVAL", interval)
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "FROMDT", fromdt)
    epqs = Replace(epqs, "UNTILDT", untildt)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = symbol
    w.[B1] = interval
    w.[C1] = fromdt
    w.[d1] = untildt

    n = o.Count     ' nombre de lignes
    p = o(1).Count  ' nombre de colonnes

    w.[A3].Resize(1, p) = o(1).Keys()
    For k = 1 To n
        w.[A4].Rows(k).Resize(1, p) = o(k).Items()
    Next k
    
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
End Sub




' *************************************
' Méthode générique à personnaliser
' *************************************
' Auteur : ** les auteurs **
' Lien vers la page API : https://rapidapi.com/...
' Params avec valeur d'exemple :
'  sheetname = " ** nom de la méthode ** " nom de la feuille crée
'  params1 = "exemple1"
'  params2 = "exemple2"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub maMethode(Optional ByVal sheetname = " ** maMethode ** ", Optional ByVal monparam1 = "exemple1", Optional ByVal monparam2 = "exemple2")
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = " ** mettre son nom d'hôte ici ** "
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/ ** son endpoint ? param1=PARAM1&param2=PARAM2" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    currencyname = WorksheetFunction.EncodeURL(currencyname)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "PARAM1", monparam1)
    epqs = Replace(epqs, "PARAM2", monparam2)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = param1
    w.[B1] = param2
    
    
    ' Puis écire toutes les informations sur la feuille w
    ' ...
    
End Sub

' ***************************
' Earning Per Share Estimate
' ***************************
' Auteur : Agathe Echegut, Axelle Bouy, Tim Hembise
' Lien vers la page API :
'   https://rapidapi.com/linuz/api/indian-stock-exchange-api2/
'
' Params avec valeur d'exemple :
'  sheetname    = "EPS_Estimates" nom de la feuille créée
'  stock_id     = "TCS"
'  measure_code = "EPS" ou CPS, CPX, DPS, EBI, EBT, GPS, GRM, NAV, NDT, NET, PRE, ROA, ROE, SAL
'  period_type  = "Annual" ou Interim
'  data_type    = "Estimates" ou Actuals
'  age          = "Current" ou OneWeekAgo, ThirtyDaysAgo, SixtyDaysAgo, NinetyDaysAgo
'
' Retour :
'  Résultats formatés sur feuille Excel [sheetname]

Public Sub EPS_Estimates( _
    Optional ByVal sheetname As String = "EPS_Estimates", _
    Optional ByVal stock_id As String = "TCS", _
    Optional ByVal measure_code As String = "EPS", _
    Optional ByVal period_type As String = "Annual", _
    Optional ByVal data_type As String = "Estimates", _
    Optional ByVal age As String = "Current" _
)
    
    'déclarations
    Dim hote As String, epqs As String, u As String, s As String
    Dim hAPI As Variant
    Dim w As Worksheet
    Dim o As Object
    Dim snaps As Object, snaps4 As Object, snap As Object
    Dim key As Variant, keysToInclude As Variant, foundmatch As Variant
    Dim col As Long, j As Long, ligne As Long
    Dim idx As Long, an As Long, annéeActuelle As Long
    Dim firstFutureIndex As Long
    Dim t As ListObject
        
    ' 1) Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "indian-stock-exchange-api2.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/stock_forecasts?stock_id=TCS&measure_code=EPS&period_type=Annual&data_type=Estimates&age=Current" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    stock_id = WorksheetFunction.EncodeURL(stock_id)
    measure_code = WorksheetFunction.EncodeURL(measure_code)
    period_type = WorksheetFunction.EncodeURL(period_type)
    data_type = WorksheetFunction.EncodeURL(data_type)
    age = WorksheetFunction.EncodeURL(age)
    
    ' Intégration des PARAMS dans la Query String: remplacement des valeurs
    epqs = Replace(epqs, "TCS", stock_id)
    epqs = Replace(epqs, "EPS", measure_code)
    epqs = Replace(epqs, "Annual", period_type)
    epqs = Replace(epqs, "Estimates", data_type)
    epqs = Replace(epqs, "Current", age)
       
    u = "https://" & hote & epqs
    hAPI = Array( _
        "X-Rapidapi-Key: " & apikey, _
        "X-Rapidapi-Host: " & hote, _
        "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    
    ' 2) Déchiffrement expression JSON en objet VBA
    On Error GoTo JsonError
    Set o = JsonConverter.ParseJson(s)
    On Error GoTo 0
    
    If Not o.Exists("periods") Then
        status = "JSON inattendu (pas de 'periods')"
        Call testErreur(resultat:=s)
        MsgBox "Réponse API inattendue (pas de 'periods').", vbCritical, "MBF Analytics"
        Exit Sub
    End If

    
    ' 3) Chargement sur feuille Excel des cartouches
    
    ' Recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    With w
        ' Purger la feuille w
        .Cells.Clear
        
        ' Inscrire les paramètres de la requête API sur la feuille w
        .[A1] = "Stock_id"
        .[A2] = "Measure_code"
        .[D1] = "Period_type"
        .[D2] = "Data_type"
        .[G1] = "Age"
        .[G2] = "Currency"
            
        .[B1] = stock_id
        .[B2] = measure_code
        .[E1] = period_type
        .[E2] = data_type
        .[H1] = age
        
    End With
                          
    ' 4) définition des champs pertinents à conserver (filtrage inclusif)
    keysToInclude = Array("Mean", "High", "Low", "Median", "StandardDeviation", "Age", "SmartEstimate")
    
    ' 5) Ecriture des titres dynamiques (firstFutureIndex comme modèle)
    ' Trouver le premier index futur comme modèle
    annéeActuelle = Year(Date)
    firstFutureIndex = 0 ' Initialiser l'index à zéro
        
    ' Recherche de la première période future
    For idx = 1 To o("periods").Count
        If o("periods")(idx).Exists("CalendarYear") Then
            an = o("periods")(idx)("CalendarYear")   ' Extrait l'année de la période de l'API
                
            If an > annéeActuelle Then
                firstFutureIndex = idx
                Exit For ' On a trouvé le premier index, on sort de la boucle
            End If
        End If
    Next idx
    
    If firstFutureIndex = 0 Then
        status = "Aucune période future trouvée"
        Call testErreur(resultat:=s)
        MsgBox "Aucune période future (CalendarYear > " & annéeActuelle & _
               ") trouvée pour " & stock_id & ".", vbInformation, "MBF Analytics"
        Exit Sub
    End If
        
    'sécurité si l'API renvoie bien des snapshots pour cette période modèle
    If o("periods")(firstFutureIndex).Exists("EstimateSnapshots") Then
        If o("periods")(firstFutureIndex)("EstimateSnapshots").Exists("EstimateSnapshot") Then
    
            'on récupère les snapshots de periods(firstFutureIndex)
            Set snaps4 = o("periods")(firstFutureIndex)("EstimateSnapshots")("EstimateSnapshot")
            If snaps4.Count = 0 Then
                MsgBox "Aucun snapshot trouvé pour la période future modèle.", vbInformation, "MBF Analytics"
                Exit Sub
            End If
            
            Set snap = snaps4(1)       'on prend le 1er snapshot comme modèle de structure
                
            col = 2 'colonne B
            For Each key In snap.Keys
                
                ' Recherche la clé (key) dans le tableau d'inclusion (keysToInclude)
                On Error Resume Next
                foundmatch = Application.Match(key, keysToInclude, 0)
                On Error GoTo 0
                
                If Not IsError(foundmatch) Then
                    w.Cells(4, col) = key 'ligne des titres
                    col = col + 1
                End If
            
            Next key
                
            ' Ecriture du titre CalendarYear
            w.Cells(4, 1) = "CalendarYear"
            
            ' Ecriture la devise une seule fois en H2
            If snap.Exists("CurrencyCode") Then
                w.[H2] = snap("CurrencyCode")
            End If
    
        End If
    End If
        
    
    ' 6) Grand tableau générique pour les périodes futures
    
    ' Initialisation de la ligne de depart excel
    ligne = 5
    For idx = 1 To o("periods").Count
               
        ' Vérifie qu'on a CalendarYear
        If o("periods")(idx).Exists("CalendarYear") Then
            an = o("periods")(idx)("CalendarYear")   ' Extrait l'année de la période de l'API
            
            ' Filtre: on ne traite et n'affiche que les données dont l'année est postérieure à l'année actuelle
            If an > annéeActuelle Then
           
               'sécurité si l'API renvoie des snapshots
               If o("periods")(idx).Exists("EstimateSnapshots") Then
                    If o("periods")(idx)("EstimateSnapshots").Exists("EstimateSnapshot") Then
            
                        Set snaps = o("periods")(idx)("EstimateSnapshots")("EstimateSnapshot")
                                            
                        For j = 1 To snaps.Count
    
                            w.Cells(ligne, 1) = an
                            
                            Set snap = snaps(j)
                            col = 2                     'on repart de la colonne B à chaque nouvelle ligne
                        
                            For Each key In snap.Keys
                                ' On gère l'erreur #N/A si la clé n'est pas trouvée dans keysToInclude
                                On Error Resume Next
                                foundmatch = Application.Match(key, keysToInclude, 0)
                                On Error GoTo 0
                                
                                ' Si Match réussit (n'est pas une erreur #N/A), on affiche la colonne
                                If Not IsError(foundmatch) Then
                                    w.Cells(ligne, col) = snap(key)
                                    col = col + 1
                                End If
                            Next key
                        
                           ' Avance à la ligne suivante pour le prochain snapshot
                           ligne = ligne + 1
                        
                        Next j  ' Fin de la boucle des snapshots
                    End If
                End If
            
            End If
        End If
        
    Next idx
    
    
    ' 7) mise en forme finale (tableau structuré)
    Set t = w.ListObjects.Add(xlSrcRange, w.[A4].CurrentRegion, xlYes)

    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
    
    Exit Sub

JsonError:
    status = "Erreur de parsing JSON"
    Call testErreur(resultat:=s)
    MsgBox "Réponse API non valide (JSON).", vbCritical, "MBF Analytics"
End Sub



' *************************************
' Cotation journalière historique d'une période donnée Yahoo Finance
' *************************************
' Auteur : Théo, Mathieu, Noé
' Lien vers la page API : https://rapidapi.com/calvinlo/api/yh-finance-complete
' Params avec valeur d'exemple :
'   sheetname = "yhf_historical" nom de la feuille crée
'   ticker    = "AAPL"           Ticker boursier (ex: AAPL, TSLA, ORA.PA)
'   sdate     = "2024-01-10"     Date de début (format YYYY-MM-DD) à partir du 1 janvier 1970
'   edate     = "2024-02-16"     Date de fin (format YYYY-MM-DD) jusqu'à j-7
' Retour
'   Feuille Excel avec historique de cotation nettoyé et formaté
Public Sub yhf_historical(Optional ByVal sheetname = "yhf_historical", Optional ByVal ticker = "AAPL", Optional ByVal sdate = "2024-01-10", Optional ByVal edate = "2024-02-16")
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "yh-finance-complete.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/yhfhistorical?ticker=TICKER&sdate=SDATE&edate=EDATE"
    
    ' Encodage URL des PARAMS (Optionnel, ticker peut contenir des caractères spéciaux)
    ' ticker = WorksheetFunction.EncodeURL(ticker)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "TICKER", ticker)
    epqs = Replace(epqs, "SDATE", sdate)
    epqs = Replace(epqs, "EDATE", edate)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur API : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Vérification qu'il y a des données
    If o.Count = 0 Then
        MsgBox "Aucune donnée.", vbExclamation
        Exit Sub
    End If

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les informations (Titre et Période)
    With w.Range("A1")
        .Value = "Historique : " & ticker
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    w.Range("A2").Value = "Période : " & sdate & " au " & edate
    
    
    ' Variables de dimension
    Dim n As Long, p As Long
    n = o.Count
    p = o(1).Count
    
    ' Récupération des clés (en-têtes)
    Dim keys As Variant
    keys = o(1).keys()
    
    ' Identification de la colonne Date pour le nettoyage
    Dim colDate As Integer
    colDate = -1
    
    Dim i As Integer
    For i = 0 To p - 1
        Dim kName As String
        kName = LCase(keys(i))
        ' On repère la colonne qui contient "date"
        If kName = "date" Or kName = "datetime" Then colDate = i
    Next i
    
    ' Écriture des En-têtes ligne 4
    w.Cells(4, 1).Resize(1, p).Value = keys
    
    ' Écriture des Données avec nettoyage de la date
    Dim startRow As Long
    startRow = 5
    
    Dim r As Long
    Dim rowItems As Variant
    Dim cleanDate As String
    
    For r = 1 To n
        ' Récupère la ligne de données sous forme de tableau (Array)
        rowItems = o(r).Items()
        
        If colDate >= 0 Then
            cleanDate = CStr(rowItems(colDate))
            
           
            cleanDate = Replace(cleanDate, "T", " ")   ' Remplace le T par un espace
            cleanDate = Replace(cleanDate, "000Z", "") ' Supprime le 000Z
          
            
            rowItems(colDate) = cleanDate
        End If
        
        ' Écriture de la ligne modifiée dans Excel
        w.Cells(startRow + r - 1, 1).Resize(1, p).Value = rowItems
    Next r


    
    Dim tblRange As Range
    Set tblRange = w.Range(w.Cells(4, 1), w.Cells(startRow + n - 1, p))
    
    ' Style En-têtes
    With w.Range(w.Cells(4, 1), w.Cells(4, p))
        .Interior.Color = RGB(52, 73, 94)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .RowHeight = 25
    End With
    
    ' Bordures du tableau
    tblRange.Borders.LineStyle = xlContinuous
    tblRange.Borders.Color = RGB(200, 200, 200)
    
    ' Formatage des Nombres et des Dates
    Dim c As Integer
    For c = 1 To p
        ' Index colonne Excel = Index Array + 1
        
        ' 1. Formatage DATE
        If (c - 1) = colDate Then
            w.Columns(c).NumberFormat = "dd/mm/yyyy hh:mm"
            w.Columns(c).HorizontalAlignment = xlRight
        
        ' 2. Formatage VOLUME (Nombre entier)
        ElseIf InStr(1, LCase(keys(c - 1)), "volume") > 0 Then
             w.Columns(c).NumberFormat = "#,##0"
             
        ' 3. Formatage PRIX (Décimales)
        Else
             w.Columns(c).NumberFormat = "#,##0.00"
        End If
    Next c
    
    ' Ajustement automatique des largeurs
    w.UsedRange.Columns.AutoFit

    ' Figer les volets
    w.Activate
    ActiveWindow.FreezePanes = False
    w.Rows("5:5").Select
    ActiveWindow.FreezePanes = True
    w.Range("A1").Select
    
End Sub



' *************************************
' Méthode générique à personnaliser
' *************************************
' Auteur : Juliette Soucasse et Lilou Droulers
' Lien vers la page API : https://rapidapi.com/contact-8Fn44oGJQx/api/financial-modeling-prep/playground/apiendpoint_29b54ea8-a339-4f64-a8e3-e9d4cda91564
' Params avec valeur d'exemple :
'  sheetname = "cashflow_statements" nom de la feuille crée
'  period = "annual"
'  limit = "50"
' symbol = "AAPL"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub financial_modeling_prep(Optional ByVal sheetname = "financial_modeling_prep", Optional ByVal period = "annual", Optional ByVal limit = "50", Optional ByVal symbol = "AAPL")
    
    ' dictionnaire financier anglais
    ' amots = Array( ... )
    
    amots = Array( _
        "acquire", "acquired", "accounts", "activities", "amortization", "and", _
        "available", "based", "benefit", "businesses", "call", "cash", "common", _
        "compensation", "contract", "customer", "decrease", "deferred", "depletion", _
        "depreciation", "debt", "dividends", "effect", "equipment", "equivalents", _
        "expense", "financing", "for", "from", "income", "including", "increase", _
        "inventories", "investing", "investments", "issuance", "liabilities", _
        "liability", "long", "loss", "maturities", "net", "noncash", "of", _
        "operating", "other", "paid", "paper", "payments", "plant", "period", _
        "prepayments", "proceeds", "property", "provided", "rate", "receivable", _
        "receivables", "repayments", "repurchase", "restricted", "sale", "sales", _
        "securities", "share", "symbol", "tax", "taxes", "to", "used", "with", _
        "withholding" _
)
    
    
    Dim key As Variant
    Dim i As Long
    Dim n As Long
    Dim p As Long
    Dim cleOriginale As String
    Dim tableauMots() As String
    Dim x As Long
    
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "financial-modeling-prep.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/v3/cash-flow-statement-as-reported/SYMBOL?period=PERIOD&limit=LIMIT&symbol=SYMBOL" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    period = WorksheetFunction.EncodeURL(period)
    limit = WorksheetFunction.EncodeURL(limit)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "PERIOD", period)
    epqs = Replace(epqs, "LIMIT", limit)
    epqs = Replace(epqs, "SYMBOL", symbol)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)


' Chargement Excel
    Set w = getOrCreateSheet(sheetname)
    w.Cells.Clear
    
    ' En-têtes paramètres
    w.[A1] = period
    w.[B1] = limit
    w.[C1] = symbol
    
  
    p = o.Count
    
    j = 1
    Set c = w.[A3]
    
    Do
        k = ""
        cleOriginale = ""
        
        ' Chercher une colonne vide
        For i = 1 To p
            If o(i).Count > 0 Then
                cleOriginale = o(i).Keys()(0)
                k = cleOriginale
                Exit For
            End If
        Next i
        
        
    If k <> "" Then
    
        ' Si k est trouvé, on l'écrit

            ' decoupe k selon dictionnaire
            ' pour chaque mot t dans amots
            ' a = split(k, t)
            ' join( a," " & t & " ")

            ' k = trim(k)
            ' k =replace(k, "  ", " ")
            
            
            
            For Each mot In amots
                ' On découpe la chaine k autour du mot trouvé
                a = Split(k, mot)
                
                ' Si le mot est trouvé (le tableau a plus d'1 élément), on recolle avec des espaces
                If UBound(a) > 0 Then
                    k = Join(a, " " & mot & " ")
                End If
            Next mot
            
            ' Supprimer les doubles espaces et espaces aux extrémités
            k = Trim(k)
            Do While InStr(k, "  ") > 0
                k = Replace(k, "  ", " ")
            Loop
            k = WorksheetFunction.Proper(k)


            tableauMots = Split(k, " ")
            k = tableauMots(0)
            
            ' On compare chaque mot avec le précédent
            For x = 1 To UBound(tableauMots)
                If LCase(tableauMots(x)) <> LCase(tableauMots(x - 1)) Then
                    k = k & " " & tableauMots(x)
                End If
            Next x
            

            c.Offset(0, j - 1) = k
            
            ' Parcourir les dictionnaires
            For i = 1 To p
                If o(i).Exists(cleOriginale) Then
                    Set ce = c.Offset(i, j - 1)
                    ce.Value = o(i)(cleOriginale)
                    
                    If cleOriginale = "date" Then
                        ce.NumberFormat = "dd mmm yyyy"
                    Else
                         ce.NumberFormat = "### ### ### ##0"
                    End If
                    
                    o(i).Remove cleOriginale
                End If
            Next i
            j = j + 1
    End If
    
        
    Loop Until k = ""
    
    Set t = w.ListObjects.Add(xlSrcRange, w.[A3].CurrentRegion, , xlYes)
    
End Sub



' *************************************
' Composantes financières d'une société
' *************************************
' Auteur : Elio Arico et Armand Geffroy
' Lien page API : https://rapidapi.com/letscrape-6bRBa3QguO5/api/real-time-finance-data/playground/endpoint_0ad54c53-f027-44ec-addd-5908aa6ea55c?env=undefined
' Params avec valeur d'exemple :
'  sheetname = "financialdata" nom de la feuille crée
'  symbol ou id = "AAPL:NASDAQ"  Apple sur la bourse du NASDAQ
'  period = "quarterly" Trimestriel et "annual" Annuel
' Retour :
'  résultats formatés sur feuille excel sheetname
'
' Exemple invocation financialdata
'    Call m.cash_flow(sheetname:="cash_flow", symbol:="AAPL:NASDAQ", period:="quarterly")
'
Public Sub cash_flow(Optional ByVal sheetname = "cashflow", Optional ByVal symbol = "AAPL:NASDAQ", Optional ByVal period = "quarterly")
    
Dim key As Variant
Dim i As Long
Dim n As Long
Dim p As Long
      
    
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Parties de la Requête de l'API avec mot générique pour les paramètres
    hote = "real-time-finance-data.p.rapidapi.com"
    epqs = "/company-cash-flow?symbol=SYMBOL&period=PERIOD"
    
    
    ' Encodage URL des PARAMS
    symbol = WorksheetFunction.EncodeURL(symbol)
    period = WorksheetFunction.EncodeURL(period)
    
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "PERIOD", period)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-RapidAPI-Key: " & apikey, "X-RapidAPI-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = "Symbol"
    w.[A2] = "Period"
    w.[D1] = "Currency"
    
    
    w.[B1] = "AAPL:NASDAQ"
    w.[B2] = "Quarterly"
    w.[E1] = "USD"
    

    'Boucle pour obtenir les donnés

    n = 5                                       ' ligne de départ
    
    Dim StartColumn As Long
    StartColumn = 1                     'colonne de départ

    Set o2 = o("data")("cash_flow")

        For i = 1 To o2.Count
        
            p = StartColumn
            For Each key In o2(i).Keys()
            
                w.[a4].Columns(p) = NomPropre(key)
                w.Cells(n, p).Value = o2(i)(key) 'valeur de la cellule ligne n colonne p
                p = p + 1
            
            Next key
    
        n = n + 1
    
    Next i
             
    'Mise en gras des titres
    
    Range("A1:A2").Font.Bold = True
    Range("D1").Font.Bold = True
    Range("A4:K4").Font.Bold = True
    
    Dim lastRow As Long, lastCol As Long
    Dim lo As ListObject

    ' Calculer la dernière ligne et colonne
    
    lastRow = w.Cells(w.Rows.Count, "A").End(xlUp).Row
    lastCol = w.Cells(4, w.Columns.Count).End(xlToLeft).Column

    ' Supprimer l'ancien tableau si il existe
    
    On Error Resume Next
    For Each lo In w.ListObjects
        lo.Delete
    Next lo
    On Error GoTo 0

    ' Créer un tableau propre
    
    Set lo = w.ListObjects.Add( _
    SourceType:=xlSrcRange, _
    Source:=w.Range(w.Cells(4, 1), w.Cells(lastRow, lastCol)), _
    XlListObjectHasHeaders:=xlYes)

    ' Nom du tableau
    
    lo.Name = "tblCashFlow"

    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
        
End Sub
    Function NomPropre(ByVal txt As String) As String
    Dim mots() As String
    Dim i As Long

    ' Remplacer les underscores par des espaces
    txt = Replace(txt, "_", " ")

    ' Mettre tout en minuscules pour uniformiser
    txt = LCase(txt)

    ' Découper en mots
    mots = Split(txt, " ")

    ' Capitaliser chaque mot
    For i = LBound(mots) To UBound(mots)
        If Len(mots(i)) > 0 Then
            mots(i) = UCase(Left(mots(i), 1)) & Mid(mots(i), 2)
        End If
    Next i

    ' Reconstituer la chaîne
    NomPropre = Join(mots, " ")
    End Function



' *************************************
' Yahoo Historical
' *************************************
' Auteur : Ornella Creignou & Anouck Nabec
' Description : This endpoint returns the history of a public stock.
' Lien vers la page API : https://rapidapi.com/belchiorarkad-FqvHs2EDOtP/api/yh-finance-complete/playground/apiendpoint_a03eb200-d01b-4980-b7ca-a3187ef5b0e3?env=undefined
' Params avec valeur d'exemple :
'  sheetname = " TSLA_HIST_TEST" nom de la feuille créée
'  ticker = "TSLA"
'  sdate = "2024-01-10" date de depart de l'historique
'  sdate = "2024-02-16" date de fin de l'historique
' Retour :
'  résultats formatés sur feuille excel sheetname

Public Sub YH_Historical(Optional ByVal sheetname = "TSLA_HIST_TEST", Optional ByVal ticker = "TSLA", Optional ByVal sdate = "2023-01-10", Optional ByVal edate = "2024-02-16")

    'déclarations
    Dim hote As String, epqs As String, u As String, s As String
    Dim hAPI As Variant
    Dim w As Worksheet
    Dim k As Long
    Dim i As Long
    
    
    ' Etat de la connexion Rapide API
    Call testErreur
     

    ' Vérifie la clé API
    Call testErreur

    ' H™te et endpoint
    hote = "yh-finance-complete.p.rapidapi.com"
    epqs = "/yhfhistorical?ticker=TSLA&sdate=2023-01-10&edate=2024-02-16"
    
    ' Encodage URL des PARAMS
    ticker = WorksheetFunction.EncodeURL(ticker)
    sdate = WorksheetFunction.EncodeURL(sdate)
    edate = WorksheetFunction.EncodeURL(edate)

    
    ' Remplacement des paramtres
    epqs = Replace(epqs, "TICKER", ticker)
    epqs = Replace(epqs, "SDATE", sdate)
    epqs = Replace(epqs, "EDATE", edate)

    ' URL finale
    u = "https://" & hote & epqs

    ' Headers API
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requte
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' DŽchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou crŽation feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
 
    ' Inscrire les paramtres de la requte API et leur valeur respectives
    w.[A1] = "ticker"
    w.[A2] = ticker

    w.[B1] = "sdate"
    w.[B2] = sdate

    w.[C1] = "edate"
    w.[C2] = edate

    If o.Count > 0 Then
        n = o.Count
        p = o(1).Count
    
    ' ƒcrire les enttes (noms des champs JSON) ˆ partir de la ligne 4
    w.[A4].Resize(1, p) = o(1).Keys()
    
    ' ƒcrire les valeurs ˆ partir de la ligne 5
    For k = 1 To n ' Boucle principale pour chaque jour de cotation.
    
    ' ƒcrire toutes les valeurs de l'enregistrement du jour
    w.[A5].Rows(k).Resize(1, p) = o(k).Items()
        
    ' Formatage des donnŽes
        ' Colonne 1 : Date/Heure (clŽ "date")
        d = w.[A5].Rows(k).Columns(1)
        
    
    ' --- Ajout des fonctions Replace ---
    
    ' 1. Remplacer "T" par un espace
    d = Replace(d, "T", " ")
    
    ' 2. Supprimer ".000Z"
    d = Replace(d, ".000Z", "")
    
    ' La variable 'd' contient maintenant la chaîne de date nettoyée
    ' Il faut la réaffecter à la cellule avant le formatage si elle a été modifiée,
    ' mais souvent on utilise la ligne suivante pour la conversion et le formatage :
    
    ' Reconversion et Formatage (si 'd' est une variable String)
    w.[A5].Rows(k).Columns(1).Value = Format(CDate(d), "dd-mm-yyyy")
    

        
        ' Colonnes 2 ˆ p : Prix et Volume
        For i = 2 To p
            ' Le volume est la 3me colonne ('volume') dans votre rŽponse JSON
            If i = 3 Then
                w.[A5].Rows(k).Columns(i).NumberFormat = "#,##0"
            Else
                w.[A5].Rows(k).Columns(i).NumberFormat = "#,##0.00"
            End If
            
        Next i
    Next k
    
    ' Mise en gras des titres et en-tte
        'Lignes 1 (A1:C1) pour les titres des paramtres
        ' Ligne 4 (A4 jusqu'ˆ la colonne p) pour les en-ttes du tableau
            w.Range("A1:C1").Font.Bold = True
            w.Range(w.Cells(4, 1), w.Cells(4, p)).Font.Bold = True
        
Dim t As ListObject

' Détermine la dernière ligne de données (Ligne 4 pour l'en-tête + n lignes de données)
Dim LastRow As Long
LastRow = 4 + n

' Création du tableau structuré (ListObject)
' Plage: de la cellule A4 (en-tête) à la cellule de la dernière colonne (p) sur la LastRow.
Set t = w.ListObjects.Add(SourceType:=xlSrcRange, Source:=w.Range(w.Cells(4, 1), w.Cells(LastRow, p)), XlListObjectHasHeaders:=xlYes)
            
Set t = w.ListObjects.Add(SourceType:=xlSrcRange, Source:=w.Range(w.Cells(1, 1), w.Cells(2, 3)), XlListObjectHasHeaders:=xlYes)
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
    End If
End Sub



****************************************
' Extraire les données des 10 cryptos majeures
' *************************************
' Auteur : Quentin Lemesnager, Tugdual Pothier, Kyan Moussavi, Baptiste Napias
' Lien vers la page API : https://rapidapi.com/tonyapi9892/api/investing-com6/playground/apiendpoint_04043732-9deb-445f-af95-d05a3680ea68
' Params avec valeur d'exemple :
'  sheetname = "CryptoMajors" nom de la feuille crée
'  params1 = "limit"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub CryptoMajors(Optional ByVal sheetname As String = "CryptoMajors ", Optional ByVal limit As Long = 10)
    
    'Declarations
    Dim epqs As String
    Dim i As Long, c As Long
    Dim hote As String
    Dim hApi As Variant
    Dim s As String
    Dim w As Worksheet
    Dim u As String
    

    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "investing-com6.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/web-crawling/api/markets/cryptocurrency/majors?limit=LIMITPARAM"
    
    ' Encodage URL des PARAMS
    limit = WorksheetFunction.EncodeURL(limit)
    
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "LIMITPARAM", limit)
        
    u = "https://" & hote & epqs
    hApi = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hApi)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscrire les paramètres de la requête API
    w.[A1] = limit
     
    
    ' Puis écire toutes les informations sur la feuille w
     c = 1
    For Each key In o(1).Keys
        w.Cells(3, c).Value = key
        If key = "PerformanceWeek" Then
            cpercent = c
            w.Columns(c).NumberFormat = "0.00%"
        End If
        w.Cells(3, c).Font.Bold = True
        c = c + 1
    Next key

    
    i = 4
    For Each crypto In o
        c = 1
        For Each key In crypto.Keys
            w.Cells(i, c).Value = crypto(key)
            If c = cpercent Then w.Cells(i, c).Value = crypto(key) / 100
            c = c + 1
        Next key
        i = i + 1
    Next crypto

    
    Set plage = w.[A3].CurrentRegion
    
   Set lo = w.ListObjects.Add(xlSrcRange, plage, , xlYes)
   
    
    
    w.UsedRange.Columns.AutoFit
    MsgBox "Crypto Majors importés"

End Sub
    
    
' *************************************
' Sentiment stock API
' *************************************
' Auteur : ** Yanis, Pierre-Alexandre, Nicolas **
' Lien vers la page API : https://rapidapi.com/kinduvarma/api/stock-sentiment-api/playground/apiendpoint_9b229260-d840-4054-88a1-ec55e893b756
' Params avec valeur d'exemple :
'  sheetname = "Sentiment_Result" nom de la feuille excel
'  ticker = "TSLA" symbol de l'action
' Retour :
'  résultats formatés sur feuille excel Sentiment_result
'
Public Sub Sentiment_Stock(Optional ByVal sheetname = "Sentiment_Resut", Optional ByVal ticker = "TSLA")
    
    ' Parties de la Requête à l'API
    ' Hote
    hote = "stock-sentiment-api.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/stock_news_sentiment/?ticker=TICKER" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    ticker = WorksheetFunction.EncodeURL(ticker)
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "TICKER", ticker)
        
    u = "https://" & hote & epqs
    
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' Déchiffrement expression JSON en objet VBA
    
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)
    
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    

   ' Inscrire les paramètres de la requête API
    w.[A1] = "Ticker"
    w.[B1] = ticker
    w.Range("A1:B1").Borders.LineStyle = xlContinuous
    
    n = o.Count     ' nombre de lignes
    p = o(1).Count  ' nombre de colonnes
    
    ' Puis écire toutes les informations sur la feuille w

    w.[A3].Resize(1, p) = o(1).Keys()
    For k = 1 To n
        w.[A4].Rows(k).Resize(1, p) = o(k).Items()
    Next k
    
    ' Ajouter les https manquant
    
    Set rg = Range("D4:D103")
    For Each cell In rg
    If Not (cell.Value Like "http*") Then
        cell.Value = "https://finviz.com" & cell.Value
    End If
    Next cell
    
    ' Faire la moyenne/max/min des notes de sentiments
    
    Set rg = Range("E4:E103")
    Moyenne = WorksheetFunction.Average(rg)
    Max = WorksheetFunction.Max(rg)
    Min = WorksheetFunction.Min(rg)
    w.[H3].Value = "Moyenne"
    w.[H4].Value = Moyenne
    w.[I3].Value = "Maximum"
    w.[I4].Value = Max
    w.[J3].Value = "Minimum"
    w.[J4].Value = Min
    w.Range("H3:J4").Borders.LineStyle = xlContinuous
    
    
    
    ' mettre sous forme de tableau
    Set rg = Range("A3:E103")
    w.ListObjects.Add(xlSrcRange, rg, , xlYes).Name = "Tableau_sentiment"
    
    ' Ajustement des colonnes automatique
    w.UsedRange.Columns.AutoFit
    
    
    

      
End Sub


' ****************************************************************
' Batterie de test des erreurs possibles du résultat des requêtes BabounSkoff
' ****************************************************************
' Paramètre
' - resultat : le résultat de la requête RapidAPI
' Retour : MsgBox puis End si une erreur est détectée
Private Sub testErreur(Optional ByVal resultat As String = "")

    If resultat Like "*" & "Enter an appropriate ticker" & "*" Then status = "Serveur API k.o."
    
    If resultat Like "*" & "not subscribed * API" & "*" Then status = "Souscription à l'API nécessaire"

    If resultat Like "*" & "Endpoint * does not exist" & "*" Then status = "Endpoint inconnue"
    
    If resultat Like "*" & "Invalid request" & "*" Then status = "Paramètres de requête incohérents"

    If resultat Like "*" & "Wrong * parameter" & "*" Then status = "Paramètres de requête incohérents"
        
    If resultat Like "*" & "data"":null" & "*" Then status = "Pas de résultats. Vérifier les paramètres"
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requêtes dépassé."
    
    ' ajouter ses propres détctions d'erreur ici et signaler l'auteur et l'objectif
    ' By XXXXX Objet YYYYY
    ' If resultat Like "*" & "??? * ???" & "*" Then status = "?????"
        

    If status <> "ok" Then
        MsgBox "Erreur : " & status & vbCrLf & vbCrLf & resultat, vbCritical, "MBF Analytics"
        End
    End If

End Sub

' *************************************
' Méthode générique à personnaliser
' *************************************
' Auteur : Selyana Luximon et Alissia Drici
' Lien vers la page API : https://rapidapi.com/principalapis/api/currency-conversion-and-exchange-rates/playground/endpoint_01c2f371-2ab0-4e56-98f4-e4f4149e9cfc
' Params avec valeur d'exemple :
'  sheetname = " ** currency_extracter** " nom de la feuille crée
'  params1 = "start_date"
'  params2 = "end_date"
'  params3 = "base"
'  params4 = "symbols"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub currency_extractor(Optional ByVal sheetname = "currency_extracter", Optional ByVal start_date = "2019-01-01", Optional ByVal end_date = "2019-01-03", Optional ByVal base = "USD", Optional ByVal symbols = "EUR%2C%20GBP")
   
    ' Parties de la Requête à l'API
    ' Hote
    hote = "currency-conversion-and-exchange-rates.p.rapidapi.com"
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/timeseries?start_date=START_DATE&end_date=END_DATE&base=BASE&symbols=SYMBOLS" ' avec mot générique pour les paramètres
   
    Debug.Print "epqs:", epqs
    symbols_non_encodee = symbols
    ' Encodage URL des PARAMS
    start_date = WorksheetFunction.EncodeURL(start_date)
    end_date = WorksheetFunction.EncodeURL(end_date)
    base = WorksheetFunction.EncodeURL(base)
    symbols = WorksheetFunction.EncodeURL(symbols)
   
    ' Etat de la connexion Rapide API
    Call testErreur
       
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "START_DATE", start_date)
    epqs = Replace(epqs, "END_DATE", end_date)
    epqs = Replace(epqs, "BASE", base)
    epqs = Replace(epqs, "SYMBOLS", symbols)
   
    Debug.Print "epqs:", epqs
   
    u = "https://" & hote & epqs
   
    Debug.Print "u:", u
   
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
   
   
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
   
    ' Batterie de test d'erreur sur resultat de la requête
    Call testErreur(resultat:=s)
   
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
       
    ' Déchiffrement expression JSON en objet VBA
   
    Dim o As Object
    Set o = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
   
    ' recherche feuille Excel ou création feuille
    Set w = getOrCreateSheet(sheetname)
   
    ' Purger la feuille
    w.Cells.Clear
   
    ' Inscrire les paramètres de la requête API
    w.[A1] = start_date
    w.[B1] = end_date
    w.[C1] = base
    w.[D1] = symbols_non_encodee
   
   
    ' Puis écire toutes les informations sur la feuille w
    ' ...
   
    ' Extraction des données JSON
    Dim rates As Object, d As Variant, curr As Variant
    Set rates = o("rates")
   
    Dim row As Long, col As Long
   
    ' Inscription des en-têtes : Date + chaque symbole
    w.Cells(3, 1).Value = "Date"
    col = 2
    For Each curr In o("rates")(o("rates").Keys()(0)).Keys
        w.Cells(3, col).Value = curr
        col = col + 1
    Next curr
   
   

   
    ' Remplissage du tableau avec les valeurs
    row = 4
    Dim valueDate As Variant, valueCurr As Variant
   
    For Each d In rates.Keys
        w.Cells(row, 1).Value = d   ' Date
       
        col = 2
        For Each valueCurr In rates(d).Keys
            w.Cells(row, col).Value = rates(d)(valueCurr)  ' Valeur du taux
            col = col + 1
        Next valueCurr
       
        row = row + 1
    Next d
   
    Set t = w.ListObjects.Add(xlSrcRange, w.[A3].CurrentRegion, , xlYes)
   
    End
   
    ' Mise en forme automatique
    w.Columns.AutoFit
    w.Activate
   
Dim lastRow As Long, lastCol As Long
    Dim chartObj As ChartObject
    Dim chartRange As Range
   
    ' Détection des bornes du tableau
    lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).row
    lastCol = w.Cells(3, w.Columns.Count).End(xlToLeft).Column

    ' Plage de données du tableau
    Set chartRange = w.Range(w.Cells(3, 1), w.Cells(lastRow, lastCol))
   
    ' Suppression d’un éventuel graphique précédent
    For Each chartObj In w.ChartObjects
        chartObj.Delete
    Next chartObj
   
    ' Création du graphique
   Set chartObj = w.ChartObjects.Add( _
    Left:=w.Cells(3, 1).Left, _
    Top:=w.Cells(lastRow + 2, 1).Top, _
    Width:=600, Height:=350)
   
   With chartObj.Chart
    .ChartType = xlLine
    .SetSourceData Source:=chartRange
    .HasTitle = True
    .ChartTitle.Text = "Évolution des taux"

    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Text = "Date"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "Taux"

    ' --- Détection automatique du min et max ---
    Dim dataRange As Range
    Dim Ymin As Double, Ymax As Double

    Set dataRange = w.Range(w.Cells(4, 2), w.Cells(lastRow, lastCol))

    Ymin = Application.WorksheetFunction.Min(dataRange)
    Ymax = Application.WorksheetFunction.Max(dataRange)

    .Axes(xlValue).MinimumScale = Ymin
    .Axes(xlValue).MaximumScale = Ymax

    .Legend.Position = xlLegendPositionBottom
    End With

Dim normStartCol As Long
normStartCol = lastCol + 2   ' on commence 2 colonnes plus loin

' Écrire les en-têtes normalisées
w.Cells(3, normStartCol).Value = "Date"
col = normStartCol + 1

For Each curr In o("rates")(o("rates").Keys()(0)).Keys
    w.Cells(3, col).Value = curr & " (norm)"
    col = col + 1
Next curr

' Récupérer les premières valeurs (pour base 100)
Dim firstValues As Object
Set firstValues = CreateObject("Scripting.Dictionary")

For Each curr In o("rates")(o("rates").Keys()(0)).Keys
    firstValues(curr) = rates(o("rates").Keys()(0))(curr)
Next curr

' Remplir les données normalisées
row = 4
For Each d In rates.Keys
    w.Cells(row, normStartCol).Value = d
   
    col = normStartCol + 1
    For Each curr In firstValues.Keys
        w.Cells(row, col).Value = (rates(d)(curr) / firstValues(curr)) * 100
        col = col + 1
    Next curr
   
    row = row + 1
Next d


' -------------------------
'   SECOND GRAPHIQUE NORMALISE
' -------------------------

Dim lastColNorm As Long
lastColNorm = w.Cells(3, w.Columns.Count).End(xlToLeft).Column

Dim chartRangeNorm As Range
Set chartRangeNorm = w.Range(w.Cells(3, normStartCol), w.Cells(lastRow, lastColNorm))

Dim chartObj2 As ChartObject
Set chartObj2 = w.ChartObjects.Add( _
    Left:=w.Cells(3, normStartCol).Left, _
    Top:=w.Cells(lastRow + 2, normStartCol).Top, _
    Width:=600, Height:=350)

With chartObj2.Chart
    .ChartType = xlLine
    .SetSourceData Source:=chartRangeNorm
    .HasTitle = True
    .ChartTitle.Text = "Évolution Normalisée des Taux (Base 100)"

    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Text = "Date"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "Indice Base 100"

    .Legend.Position = xlLegendPositionBottom
End With
End Sub




' *************************************
' Importation d'options  
' *************************************
' Auteur : Vanessa JIN, Phuong Linh NHU, Agathe CANEL
' Lien vers la page API : https://rapidapi.com/3b-data-3b-data-default/api/yahoo-finance-real-time1/playground/apiendpoint_00691d4c-5c21-4972-8560-4c03880ebc7f

' Params avec valeur d'exemple :
' symbol = "WOOF"
'  langue = "en-US"
'  region = "US"

' Retour: Feuille "Options" : infos du sous-jacent + dates d'expiration + tableaux des calls puis des puts

Private Sub testeErreur(Optional ByVal resultat As String = "")
    
    If resultat Like "*" & "not subscribed * API" & "*" Then status = "Souscription à l'API necessaire"

    If resultat Like "*" & "Endpoint * does not exist" & "*" Then status = "Endpoint inconnue"
    
    If resultat Like "*" & "Invalid request" & "*" Then status = "Parametres de requete incoherents"

    If resultat Like "*" & "Wrong * parameter" & "*" Then status = "Parametres de requete incoherents"
        
    If resultat Like "*" & "data"":null" & "*" Then status = "Pas de resultats. Verifier les parametres"
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requetes depasse."
    
    ' ajouter ses propres detections d'erreurs ici et signaler l'auteur et l'objectif
    ' By XXXXX Objet YYYYY
    ' If resultat Like "*" & "??? * ???" & "*" Then status = "?????"
        
    If status <> "ok" Then
        MsgBox "Erreur : " & status & vbCrLf & vbCrLf & resultat, vbCritical, "MBF Analytics"
        End
    End If

End Sub


Public Sub Options(Optional ByVal sheetname = "Options", Optional ByVal symbol = "NVDA", Optional ByVal langue = "en-US", Optional ByVal region = "US")
    
    Dim hote As String, epqs As String, u As String
    Dim hAPI As Variant, s As String
    Dim o As Object
    

    ' 1- Parties de la requete l'API
    ' --------------------------------------------------------------
    ' Hote RapidAPI
    hote = "yahoo-finance-real-time1.p.rapidapi.com"
    ' Endpoint + QueryString generique
    epqs = "/stock/get-options?symbol=SYMBOL&lang=LANGUE&region=REGION"


    ' 2- Encodage URL des parametres utilisateur
    ' --------------------------------------------------------------
    symbol = WorksheetFunction.EncodeURL(symbol)
    langue = WorksheetFunction.EncodeURL(langue)
    region = WorksheetFunction.EncodeURL(region)
    
    
    ' 3- Integration des parametres dans la Query String + appel HTTP + gestion des erreurs
    ' --------------------------------------------------------------
    ' Reinitialise le statut / detecte les erreurs de cle par defaut
    Call testeErreur
        
    ' Integration des PARAMS dans la Query String
    epqs = Replace(epqs, "SYMBOL", symbol)
    epqs = Replace(epqs, "LANGUE", langue)
    epqs = Replace(epqs, "REGION", region)
    
    ' URL complete
    u = "https://" & hote & epqs
    ' Headers RapidAPI
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    ' Requete HTTP
    s = h.getHttp(url:=u, headers:=hAPI)

    ' Affiche la reponse brute dans la fenetre Execution
    Debug.Print s
    
    ' Test d'erreur
    Call testeErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
        
    ' 4 - Dechiffrement expression JSON en objet VBA
    ' --------------------------------------------------------------
    Set o = JsonConverter.ParseJson(s)
    
    ' Verifier si les variables sont bien ecrites
    Dim resultCount As Long
    resultCount = o("optionChain")("result").Count
    
    If resultCount = 0 Then
        Dim msg As String
        msg = "Parametres invalides detectes :" & vbCrLf
        msg = msg & "Symbole : " & symbol & vbCrLf
        msg = msg & "Langue : " & langue & vbCrLf
        msg = msg & "Region : " & region & vbCrLf
        msg = msg & vbCrLf & "Verifiez ces valeurs et reessayez."
        MsgBox msg, vbCritical, "Erreur API Yahoo Finance"
        Exit Sub
    End If
    
    ' Structure du JSON de l'AP
    Dim chain As Object, quote As Object
    Set chain = o("optionChain")("result")(1)
    Set quote = o("optionChain")("result")(1)("quote")


    ' 5 - Création du tableau des informations du sous-jacents
    ' --------------------------------------------------------------

    Dim wsOptions As Worksheet
    Set wsOptions = getOrCreateSheet(sheetname)  ' cree ou recupere la feuille
    wsOptions.Cells.Clear                                          ' purge du contenu

    'Le cas ou l'entreprise choisie par l'utilisateur ne possede pas d'option
    If chain("options").Count = 0 Then
        wsOptions.Range("A1").value = "Cette entreprise n'a pas d'option associée"
        wsOptions.Range("A1").Font.Color = vbRed
        wsOptions.Range("A1").Font.Bold = True
    'Le cas ou il y a des options
    Else
        ' Ligne d'en-tetes du tableau
        wsOptions.Range("A1:I1").value = Array("Symbol", "Company", "Exchange", "Currency", "Price", "Change %", "Volume", "Market Cap", "Expiration Dates")
    
        ' Mise en forme de l'en-tetes
        wsOptions.Rows(1).Font.Bold = True
        wsOptions.Range(wsOptions.Cells(1, 1), wsOptions.Cells(1, 9)).Interior.Color = RGB(230, 230, 230)
        
        
        ' Infos du sous-jacent
        wsOptions.Cells(2, 1).value = chain("underlyingSymbol")
        wsOptions.Cells(2, 2).value = quote("longName")
        wsOptions.Cells(2, 3).value = quote("exchange")
        wsOptions.Cells(2, 4).value = quote("currency")
        wsOptions.Cells(2, 5).value = quote("regularMarketPrice")
        wsOptions.Cells(2, 6).value = quote("regularMarketChangePercent")
        wsOptions.Cells(2, 7).value = quote("regularMarketVolume")
        wsOptions.Cells(2, 8).value = quote("marketCap")
        
        ' Liste des dates d'expiration en colonne I
        Dim expDates As Object
        Dim k As Long, row As Long
        Set expDates = chain("expirationDates")
        row = 2
        ' Conversion en date Excel
        For k = 1 To expDates.Count
            wsOptions.Cells(row, 9).value = DateAdd("s", CLng(expDates(k)), #1/1/1970#)
            row = row + 1
        Next k
        ' Mettre la date au bon format
        wsOptions.Cells(row, 9).NumberFormat = "yyyy-mm-dd"
        
        ' Mise en forme rapide
        With wsOptions
            .Columns.AutoFit
            .Cells.HorizontalAlignment = xlCenter
        End With
    
    
        Dim opt As Object, calls As Object, puts As Object
        Dim r As Long, i As Long
        Dim hdr As Variant
    
        Set opt = chain("options")(1)
        
        Set calls = opt("calls")
        Set puts = opt("puts")
     
        
    ' 6 - Création du tableau des informations des CALLS
    ' --------------------------------------------------------------
        r = row + 2
        
        wsOptions.Cells(r, 1).value = "Call"
        wsOptions.Cells(r, 1).Font.Bold = True
        
        ' Noms en-tetes des colonnes
        Dim headers As Variant
        Dim c As Long

        headers = Array("Contract Symbol", "Strike", "Currency", "Last Price", "Change", "Percent Change", "Volume", "Bid", "Ask", "Expiration Date", "Last Trade Date", "Implied Volatility")

        For c = LBound(headers) To UBound(headers)
             wsOptions.Cells(r, c + 1).value = headers(c)
        Next c
 
        ' Mise en forme de l'en-tete
        wsOptions.Rows(r).Font.Bold = True
        wsOptions.Range(wsOptions.Cells(r, 1), wsOptions.Cells(r, UBound(headers) + 1)).Interior.Color = RGB(230, 230, 230)
        
        r = r + 1 ' premiere ligne de donnees
        
        
        Dim keysToInclude As Variant
        keysToInclude = Array("contractSymbol", "strike", "currency", "lastPrice", "change", "percentChange", "volume", "bid", "ask", "expiration", "lastTradeDate", "impliedVolatility")
        
        Dim snap As Object
        Dim key As Variant
        Dim col As Long
        Dim found As Variant
        
        ' Pour chaque CALL
        For i = 1 To calls.Count
            Set snap = calls(i)   ' un seul call
            col = 1               ' on repart colonne 1 à chaque ligne
            ' On boucle sur toutes les clés disponibles dans ce CALL
            For Each key In snap.Keys
                ' Vérifier si la clé fait partie de keysToInclude
                On Error Resume Next
                found = Application.Match(key, keysToInclude, 0)
                On Error GoTo 0
                If Not IsError(found) Then
                    ' Conversion automatique pour les timestamps
                    If key = "expiration" Or key = "lastTradeDate" Then
                        wsOptions.Cells(r, col).value = DateAdd("s", CLng(snap(key)), #1/1/1970#)
                        wsOptions.Cells(r, col).NumberFormat = "yyyy/mm/dd"
                    ' Format pour l'implied volatility
                    ElseIf key = "impliedVolatility" Then
                        wsOptions.Cells(r, col).NumberFormat = "0.00"
                        wsOptions.Cells(r, col).value = snap(key)
                    ' Tous les autres champs
                    Else
                        wsOptions.Cells(r, col).value = snap(key)
                    End If
                    col = col + 1
                End If
            Next key
            ' passer à la prochaine ligne du tableau
            r = r + 1
        Next i
            
            
    ' 7 - Création du tableau des informations des PUTS
    ' --------------------------------------------------------------
        r = r + 2  ' petite separation visuelle des 2 tableaux
        wsOptions.Cells(r, 1).value = "Put"
        wsOptions.Cells(r, 1).Font.Bold = True
                
        r = r + 1  ' Ligne des en-tetes (Puts)
                
        ' Noms en-tetes des colonnes
        For c = LBound(headers) To UBound(headers)
                wsOptions.Cells(r, c + 1).value = headers(c)
        Next c
         
        ' Mise en forme de l'en-tete
        wsOptions.Rows(r).Font.Bold = True
        wsOptions.Range(wsOptions.Cells(r, 1), wsOptions.Cells(r, UBound(headers) + 1)).Interior.Color = RGB(230, 230, 230)
            
        r = r + 1 ' premiere ligne de donnees
            
        
        Dim snapP As Object
        Dim keyP As Variant
        Dim foundP As Variant
        
        ' Pour chaque PUT
        For i = 1 To puts.Count
                Set snapP = puts(i)   ' un seul call
                col = 1               ' on repart colonne 1 à chaque ligne
                ' On boucle sur toutes les clés disponibles dans ce put
                For Each keyP In snapP.Keys
                    ' Vérifier si la clé fait partie de keysToInclude
                    On Error Resume Next
                    foundP = Application.Match(keyP, keysToInclude, 0)
                    On Error GoTo 0
                    If Not IsError(foundP) Then
                        ' Conversion automatique pour les timestamps
                        If keyP = "expiration" Or keyP = "lastTradeDate" Then
                            wsOptions.Cells(r, col).value = DateAdd("s", CLng(snapP(keyP)), #1/1/1970#)
                            wsOptions.Cells(r, col).NumberFormat = "yyyy/mm/dd"
                        ' Format pour l'implied volatility
                        ' ElseIf keyP = "impliedVolatility" Then
                            wsOptions.Cells(r, col).NumberFormat = "0.00"
                            wsOptions.Cells(r, col).value = snapP(keyP)
                        ' Tous les autres champs
                        Else
                            wsOptions.Cells(r, col).value = snapP(keyP)
                        End If
                        col = col + 1
                    End If
                 Next keyP
                 r = r + 1
        Next i
            
    End If
    
    'Mise en forme Options'
     With wsOptions
        .Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter
    End With
      
End Sub

   
       
       


' **************************************************
' MBF Analytics
' Banque de mÈthodes de DonnÈes financiËres en ligne
' Simple utilisation
' Robuste
' Personnalisable API Key
' ***************************************************


' References :  Outils > References
' 1- Microsoft HTML Object Library
' 2- Microsoft WinHTTP Services, version 5.1

' Classe
' 1- Webtool
' 2- Dictionnary


' Constantes des la classe
Private Const keylen As Integer = 50

' PropriÈtÈs de la classe
Private apikey As String    ' code perso pour se connecter ‡ Rapid API
Private h As Object         ' h objet pour les requÍtes HTTP
Public status As String     ' Ètat de la connection avec Rapid API :
                            ' ok : tout va bien
                            ' API Key invalide : problËme de clÈ API
                            ' Compte ÈpuisÈ : plus de requÍtes possible
                            ' ParamËtres de requÍte incohÈrents : Invalid request
                            ' Endpoint Inconnue : Endpoint * does not exist
                            ' Erreur non reconnue
                            

' ******************
' MÈthodes Internes
' ******************
' Constructeur de la classe
Private Sub Class_Initialize()
    Set h = New WebTool
    status = "ok"
End Sub

Public Sub initKey(key As String)
    
    apikey = Trim(key)
    
    If Len(apikey) <> keylen Then
        status = "API Key invalide"
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        End
    Else
        status = "ok"
    End If
        
    Set h = New WebTool

End Sub

' ------------------------------------------------------------
' Fonction de recherche ou crÈation d'une feuille Excel nommÈe
' FigÈe
' ParamËtre
' - sheetname : nom de la feuille Excel recherchÈe
' Retour : objet feuille Excel
Private Function getOrCreateSheet(ByVal sheetname As String) As Worksheet
    Dim w As Worksheet
    
    Set w = Nothing
    
    For Each w In Worksheets
        If w.Name = sheetname Then Exit For
    Next w
    
    If w Is Nothing Then
        Set w = Worksheets.Add
        w.Name = sheetname
    End If

    Set getOrCreateSheet = w
End Function


' ***********************************************
' ***********************************************
' Collection de MÈthodes financiËres de la classe
' ***********************************************
' ***********************************************

' ****************************************************************
' Batterie de test des erreurs possibles du rÈsultat des requÍtes
' ****************************************************************
' ParamËtre
' - resultat : le rÈsultat de la requÍte RapidAPI
' Retour : MsgBox puis End si une erreur est dÈtectÈe
Private Sub testErreur(Optional ByVal resultat As String = "")
    
    If resultat Like "*" & "not subscribed * API" & "*" Then status = "Souscription ‡ l'API nÈcessaire"

    If resultat Like "*" & "Endpoint * does not exist" & "*" Then status = "Endpoint inconnue"
    
    If resultat Like "*" & "Invalid request" & "*" Then status = "ParamËtres de requÍte incohÈrents"

    If resultat Like "*" & "Wrong * parameter" & "*" Then status = "ParamËtres de requÍte incohÈrents"
        
    If resultat Like "*" & "data"":null" & "*" Then status = "Pas de rÈsultats. VÈrifier les paramËtres"
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requÍtes dÈpassÈ."
    
    If resultat Like "*" & "have exceeded * quota" & "*" Then status = "Quota de requÍtes dÈpassÈ."
    
    If resultat Like "*" & "not valid value from parameter" & "*" Then status = "ParamËtres de requÍte incohÈrents"
    
    
    ' Erreur ajoutÈe
        
    If resultat Like "*" & "code: 500" & "*" Then status = "Serveur indisponible"
    
    ' Erreurs liÈes ‡ líauthentification
If resultat Like "*" & "401 Unauthorized*" & "*" Then status = "API Key non valide ou expirÈe"
If resultat Like "*" & "403 Forbidden*" & "*" Then status = "AccËs refusÈ : permissions insuffisantes"

' Erreurs de format ou de syntaxe
If resultat Like "*" & "400 Bad Request*" & "*" Then status = "RequÍte mal formÈe ou paramËtre invalide"
If resultat Like "*" & "invalid date*" & "*" Then status = "ParamËtre date invalide"

' Erreurs liÈes ‡ la limite de requÍtes
If resultat Like "*" & "rate limit exceeded*" & "*" Then status = "Limite de requÍtes atteinte. Attendez avant de relancer"
If resultat Like "*" & "429 Too Many Requests*" & "*" Then status = "Trop de requÍtes envoyÈes en peu de temps"

' Erreurs liÈes aux donnÈes manquantes
If resultat Like "*" & "No data found*" & "*" Then status = "Aucune donnÈe disponible pour ces paramËtres"
If resultat Like "*" & "data not available*" & "*" Then status = "DonnÈes non disponibles pour ce pays ou cette pÈriode"

' Erreurs serveur
If resultat Like "*" & "502 Bad Gateway*" & "*" Then status = "Erreur serveur temporaire (502)"
If resultat Like "*" & "503 Service Unavailable*" & "*" Then status = "Service temporairement indisponible (503)"
If resultat Like "*" & "504 Gateway Timeout*" & "*" Then status = "DÈlai d'attente dÈpassÈ, serveur lent (504)"

' Cas spÈcifiques RapidAPI
If resultat Like "*" & "You have reached your subscription limit*" & "*" Then status = "Limite de souscription RapidAPI atteinte"
If resultat Like "*" & "API key missing*" & "*" Then status = "ClÈ API manquante ou non transmise"

    
    If status <> "ok" Then
        MsgBox "Erreur : " & status & vbCrLf & vbCrLf & resultat, vbCritical, "MBF Analytics"
        End
    End If

End Sub


' *************************************
' MÈthode gÈnÈrique ‡ personnaliser -> Calendrier Èconomique via API
' *************************************
' Auteur : Valentin Grange & Marion Neveu-Perret
' Lien vers la page API : https://rapidapi.com/yasimpratama88/api/economic-calendar-api

' Params avec valeur d'exemple :
'  sheetname = "Calendrier" nom de la feuille crÈe
'  endDate = "2025-07-30" date de fin doit Ítre spÈcifiÈe en format ISO : YYYY-MM-DD et doit se situer dans la plage de dates correspondant ‡ la semaine en cours
'  countryCode = "US" Permet de filtrer les donnÈes par code de pays (voir annexe en fin de page pour la liste exhaustive). Il est possible de spÈcifier un code unique de pays pour obtenir les donnÈes d'un pays spÈcique ou de ne rien renseigner pour obtenir les donnÈes de tous les pays disponibles. Impossible de rentrer une liste de pays.
'  limit  = "500" Permet de dÈfinir le nombre maximum de lignes renvoyÈes par le tableau.
'  startDate =  "2025-07-30" date de dÈbut doit Ítre spÈcifiÈe en format ISO : YYYY-MM-DD et doit se situer dans la plage de dates correspondant ‡ la semaine en cours
' timezone = "GMT+1"  paramËtre permet de spÈcifier le fuseau horaire pour les heures des ÈvÈnements du calendrier (voir annexe de fin de page pour la liste exhaustive). Par exemple, "GMT+1" correspond ‡ l'heure de l'Europe centrale. Si un autre fuseau horaire est nÈcessaire, vous pouvez le spÈcifier sous la forme "GMT+X" ou "GMT-X", o˘ X reprÈsente l'offset horaire et X doit Ítre un nombre entier.
' volatility = "MEDIUM" permet de spÈcifier le niveau de volatilitÈ attendu pour les ÈvÈnements, avec les options "LOW" pour faible, "MEDIUM" pour modÈrÈ, "HIGH" pour ÈlevÈ, ou "NONE" pour inclure toutes les donnÈes, indÈpendamment de la volatilitÈ.
' Retour :
'  rÈsultats formatÈs sur feuille excel sheetname

'
Public Sub MacroCalendar(Optional ByVal sheetname As String = "Calendrier_1", Optional ByVal startDate As String = "2025-07-29", Optional ByVal endDate As String = "2025-07-30", Optional ByVal countryCode As String = "US,UK,DE", Optional ByVal timezone As String = "GMT+1", Optional ByVal volatility As String = "NONE", Optional ByVal limit As String = "500")
    
    ' Parties de la RequÍte ‡ l'API
    ' Hote
    hote = "economic-calendar-api.p.rapidapi.com"
    ' Endpoint?Query String avec ParamËtres gÈnÈriques en majuscule
    epqs = "/calendar?endDate=ENDDATE&countryCode=COUNTRYCODE&limit=LIMIT&startDate=STARTDATE&timezone=TIMEZONE"
    
     'Ajout pour gÈrer le paramËtre volatility avec un if
    If UCase(volatility) <> "NONE" Then
        ' Si LOW / MEDIUM / HIGH ? on l'ajoute
        epqs = "/calendar?endDate=ENDDATE&countryCode=COUNTRYCODE&limit=LIMIT&startDate=STARTDATE&timezone=TIMEZONE&volatility=VOLATILITY"
    Else
        ' Si NONE ? on ne touche pas ‡ epqs (donc pas de volatilitÈ dans líURL)
    End If
    
    'RÈcupÈration de la valeur non encodÈe de TIMEZONE pour l'inclure dans le tableau rÈsumer sous le bon format
    timezone_raw = timezone
    
    ' Encodage URL des PARAMS
    endDate = WorksheetFunction.EncodeURL(endDate)
    countryCode = WorksheetFunction.EncodeURL(countryCode)
    limit = WorksheetFunction.EncodeURL(limit)
    startDate = WorksheetFunction.EncodeURL(startDate)
    timezone = WorksheetFunction.EncodeURL(timezone)
    volatility = WorksheetFunction.EncodeURL(volatility)
    
    
    ' Etat de la connexion Rapide API
    Call testErreur
        
    ' IntÈgration des PARAMS dans la Query String
    epqs = Replace(epqs, "ENDDATE", endDate)
    epqs = Replace(epqs, "COUNTRYCODE", countryCode)
    epqs = Replace(epqs, "LIMIT", limit)
    epqs = Replace(epqs, "STARTDATE", startDate)
    epqs = Replace(epqs, "TIMEZONE", timezone)
    epqs = Replace(epqs, "VOLATILITY", volatility)
        
    u = "https://" & hote & epqs
    hAPI = Array("X-Rapidapi-Key: " & apikey, "X-Rapidapi-Host: " & hote, "Content-Type: Application/json; charset=utf-8")
    
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur resultat de la requÍte
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
    ' DÈchiffrement expression JSON en objet VBA
    
    Dim JsonC As Object
    Set JsonC = JsonConverter.ParseJson(s)

    ' Chargement sur feuille Excel
    
    ' recherche feuille Excel ou crÈation feuille
    Set w = getOrCreateSheet(sheetname)
    
    ' Purger la feuille
    w.Cells.Clear
    
    ' Inscription des intitulÈs et valeurs du tableau rÈsumÈ
    w.[J1] = "Date de dÈbut"
    w.[J2] = "Date de fin"
    w.[J3] = "Code de pays"
    w.[J4] = "Fuseau horaire"
    w.[J5] = "VolatilitÈ attendue"
    w.[J6] = "Limite de ligne"
    
    w.[K1] = startDate
    w.[K2] = endDate
    w.[K3] = countryCode
    w.[K4] = timezone_raw
    w.[K5] = volatility
    w.[K6] = limit
    
    ' Inscription des intitulÈs du tableau Èconomique
    w.Range("A1").Value = "Date"
    w.Range("B1").Value = "Time"
    w.Range("C1").Value = "Pays"
    w.Range("D1").Value = "Event"
    w.Range("E1").Value = "Volatility"
    w.Range("F1").Value = "Actual"
    w.Range("G1").Value = "Consensus"
    w.Range("H1").Value = "Previous"

    ' Puis Ècire toutes les informations sur la feuille w

' --- DÈbut : Ècrire les events JSON sous les en-tÍtes A8:G8 ---
Dim dataArr As Object
Dim itm As Object
Dim writeRow As Long

Dim dateStr As String
Dim yearN As Long, monthN As Long, dayN As Long
Dim hourN As Long, minN As Long, secN As Long
Dim dt As Date
Dim tzOffset As Long
Dim signPos As Long

' VÈrifier que JsonC contient "data"
If JsonC Is Nothing Then
    MsgBox "JSON non parsÈ (JsonC = Nothing).", vbCritical
    Exit Sub
End If

If Not JsonC.Exists("data") Then
    MsgBox "Le JSON ne contient pas la clÈ 'data'.", vbExclamation
    Exit Sub
End If

Set dataArr = JsonC("data")

' DÈterminer l'offset du fuseau timezone (ex: "GMT+1" -> 1, "GMT-2" -> -2)
tzOffset = 0
If Len(timezone) > 0 Then
    ' On suppose timezone comme "GMT+1" ou "GMT-2" ou "GMT+0"
    signPos = InStr(1, timezone, "GMT", vbTextCompare)
    If signPos > 0 Then
        On Error Resume Next
        Dim rawOffset As String
        rawOffset = Replace(timezone, "GMT", "", 1, -1, vbTextCompare)
        rawOffset = Trim(rawOffset)
        If rawOffset <> "" Then
            tzOffset = CLng(rawOffset) ' convertit "+1" ou "-2" en nombre
        Else
            tzOffset = 0
        End If
        On Error GoTo 0
    End If
End If

' Ligne de dÈpart : la premiËre ligne de donnÈes est la ligne 9 (tes en-tÍtes sur la ligne 8)
writeRow = 2

' Boucle sur les ÈlÈments
If Not dataArr Is Nothing Then
    For Each itm In dataArr
        ' 1) Date & Time : partir de dateUtc si prÈsent
        dateStr = ""
        dt = 0
        If itm.Exists("dateUtc") Then
            dateStr = itm("dateUtc") ' ex: "2025-11-17T17:30:00.000Z"
            If Len(dateStr) >= 19 Then
                ' extraire composants (on ignore millisecondes et Z)
                On Error Resume Next
                yearN = CLng(Mid$(dateStr, 1, 4))
                monthN = CLng(Mid$(dateStr, 6, 2))
                dayN = CLng(Mid$(dateStr, 9, 2))
                hourN = CLng(Mid$(dateStr, 12, 2))
                minN = CLng(Mid$(dateStr, 15, 2))
                secN = CLng(Mid$(dateStr, 18, 2))
                On Error GoTo 0
                If yearN > 1900 Then
                    dt = DateSerial(yearN, monthN, dayN) + TimeSerial(hourN, minN, secN)
                    ' Appliquer l'offset du fuseau (dateUtc -> timezone local)
                    If tzOffset <> 0 Then
                        dt = DateAdd("h", tzOffset, dt)
                    End If
                End If
            End If
        End If

        ' Ecrire colonne Date (A) et Date(B) :
        If dt <> 0 Then
            w.Cells(writeRow, "A").Value = Format(dt, "yyyy-mm-dd")
            w.Cells(writeRow, "B").Value = Format(dt, "HH:nn")
        Else
            w.Cells(writeRow, "A").Value = ""
            w.Cells(writeRow, "B").Value = ""
        End If


        ' Colonne Event (C) -> CountryCode
        If itm.Exists("countryCode") Then
            If IsNull(itm("countryCode")) Then
                w.Cells(writeRow, "C").Value = ""
            Else
                w.Cells(writeRow, "C").Value = itm("countryCode")
            End If
        Else
            w.Cells(writeRow, "C").Value = ""
        End If


        ' Colonne Event (D) -> name
        If itm.Exists("name") Then
            If IsNull(itm("name")) Then
                w.Cells(writeRow, "D").Value = ""
            Else
                w.Cells(writeRow, "D").Value = itm("name")
            End If
        Else
            w.Cells(writeRow, "D").Value = ""
        End If

        ' Colonne Volatility (E) -> volatility (ex: "LOW")
        If itm.Exists("volatility") Then
            If IsNull(itm("volatility")) Then
                w.Cells(writeRow, "E").Value = ""
            Else
                w.Cells(writeRow, "E").Value = itm("volatility")
            End If
        Else
            w.Cells(writeRow, "E").Value = ""
        End If

        ' Colonne Actual (F) -> actual
        If itm.Exists("actual") Then
            If IsNull(itm("actual")) Then
                w.Cells(writeRow, "F").Value = ""
            Else
                w.Cells(writeRow, "F").Value = itm("actual")
            End If
        Else
            w.Cells(writeRow, "F").Value = ""
        End If

        ' Colonne Consensus (G) -> consensus
        If itm.Exists("consensus") Then
            If IsNull(itm("consensus")) Then
                w.Cells(writeRow, "G").Value = ""
            Else
                w.Cells(writeRow, "G").Value = itm("consensus")
            End If
        Else
            w.Cells(writeRow, "G").Value = ""
        End If

        ' Colonne Previous (H) -> previous
        If itm.Exists("previous") Then
            If IsNull(itm("previous")) Then
                w.Cells(writeRow, "H").Value = ""
            Else
                w.Cells(writeRow, "H").Value = itm("previous")
            End If
        Else
            w.Cells(writeRow, "H").Value = ""
        End If

        writeRow = writeRow + 1
    Next itm
Else
    MsgBox "La liste data est vide.", vbInformation
End If

' --- Fin : Ècriture des events ---


' Comptabilisation du nombre de lignes renvoyÈes par l'API

w.[J7] = "Nombre de ligne"
If Not dataArr Is Nothing Then
    w.Range("K7").Value = dataArr.Count
Else
    w.Range("K7").Value = 0
End If


'***** MISE EN FORME****


' === Transforme A:G en Table Excel ===
Dim lastRow As Long
Dim lo As ListObject

' Trouver la derniËre ligne remplie dans la colonne A
lastRow = w.Cells(w.Rows.Count, "A").End(xlUp).Row

' Supprimer líancienne table si elle existe dÈj‡
On Error Resume Next
w.ListObjects("TableCalendar").Delete
On Error GoTo 0

' CrÈer la table
Set lo = w.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=w.Range("A1:H" & lastRow), _
            XlListObjectHasHeaders:=xlYes)

' Nom de la table
lo.Name = "TableCalendar"

' Style de tableau (moderne et clair)
lo.TableStyle = "TableStyleMedium9"



' === Mise en couleur simple de la colonne Volatility (colonne D) ===
Dim rngVol As Range
Dim cell As Range

Set lo = w.ListObjects("TableCalendar")
Set rngVol = lo.ListColumns("Volatility").DataBodyRange  ' Utiliser lo et non var

' Parcours des cellules
For Each cell In rngVol
    Select Case UCase(cell.Value)
        Case "LOW"
            cell.Interior.Color = RGB(255, 255, 153) ' jaune clair
        Case "MEDIUM"
            cell.Interior.Color = RGB(255, 204, 102) ' orange clair
        Case "HIGH"
            cell.Interior.Color = RGB(255, 102, 102) ' rouge clair
        Case Else
            cell.Interior.ColorIndex = xlNone ' pas de couleur
    End Select
Next cell



' === Fin de líÈcriture des donnÈes JSON : ajout des informations du tableau rÈsumÈ ===

' Nombre de lignes rÈellement retournÈes
w.[J7] = "Nombre de lignes"
w.[K7] = writeRow - 1

' Statut API
w.[J8] = "Statut API"
w.[K8] = status


' Heure de la requÍte
w.[J9] = "RequÍte effectuÈe le"
w.[K9] = Now


w.Columns("K").HorizontalAlignment = xlLeft


' === Coloration des cellules Actual (E) selon comparaison avec Consensus (F) ===
Dim iconRange As Range
Dim cell2 As Range

' DÈfinir la plage de la colonne Actual (E)
Set iconRange = w.Range("F2:F" & writeRow - 1) ' writeRow pointe sur la ligne suivante aprËs la derniËre

' Parcours des cellules
For Each cell2 In iconRange
    ' VÈrifier si Actual et Consensus ne sont pas vides et sont numÈriques
    If Len(cell2.Value) > 0 And Len(cell2.Offset(0, 1).Value) > 0 Then
        If IsNumeric(cell2.Value) And IsNumeric(cell2.Offset(0, 1).Value) Then
            If cell2.Value > cell2.Offset(0, 1).Value Then
                ' Actual > Consensus : vert
                cell2.Interior.Color = RGB(0, 176, 80)
            ElseIf cell2.Value < cell2.Offset(0, 1).Value Then
                ' Actual < Consensus : rouge
                cell2.Interior.Color = RGB(192, 0, 0)
            Else
                ' Actual = Consensus : pas de couleur
                cell2.Interior.ColorIndex = xlNone
            End If
        Else
            ' Si pas numÈrique : pas de couleur
            cell2.Interior.ColorIndex = xlNone
        End If
    Else
        ' Si une des deux cellules est vide : pas de couleur
        cell2.Interior.ColorIndex = xlNone
    End If
Next cell2



' Ajustement automatique de la largeur des colonnes
On Error Resume Next    ' …vite les erreurs si une colonne est vide
w.Cells.EntireColumn.AutoFit
On Error GoTo 0

End Sub


' *************************************
' Méthode générique à personnaliser
' *************************************
' Auteur : Juliette Morignot,  Alya Bennys
' Lien vers la page API : https://rapidapi.com/alphawave/api/financial-statements/playground/apiendpoint_b4be9cff-500e-4f64-a5a9-a8d73f9cde7a
' Params avec valeur d'exemple :
'  sheetname = "Income_statement" nom de la feuille créée
'  ticker = "AAPL"
' Retour :
'  résultats formatés sur feuille excel sheetname
'
Public Sub Income_statement( _
    Optional ByVal sheetname As String = "Income_statement", _
    Optional ByVal ticker As String = "AAPL")

    ' === Déclaration des variables ===
    Dim w As Worksheet
    Dim o As Object
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim key As Variant

    ' Parties de la Requête à l'API
    ' Hôte
    hote = "financial-statements.p.rapidapi.com"
    
    ' Endpoint?Query String avec Paramètres génériques en majuscule
    epqs = "/api/v1/resources/income-statement?ticker=TICKER" ' avec mot générique pour les paramètres
    
    ' Encodage URL des PARAMS
    ticker = WorksheetFunction.EncodeURL(ticker)
    
    ' Etat de la connexion Rapide API
    Call testErreur
    
    ' Intégration des PARAMS dans la Query String
    epqs = Replace(epqs, "TICKER", ticker)
       
    u = "https://" & hote & epqs
    
    hAPI = Array( _
        "X-Rapidapi-Key: " & apikey, _
        "X-Rapidapi-Host: " & hote, _
        "Content-Type: Application/json; charset=utf-8" _
    )
    
    ' Appel HTTP
    s = h.getHttp(url:=u, headers:=hAPI)
 
    Debug.Print s
    
    ' Batterie de test d'erreur sur résultat de la requête
    Call testErreur(resultat:=s)
    
    If status <> "ok" Then
        MsgBox "Erreur : " & status, vbCritical, "MBF Analytics"
        Exit Sub
    End If
        
 ' --- 1) Déchiffrer le JSON en objet VBA ---
Dim o As Object
Set o = JsonConverter.ParseJson(s)

' --- 2) Récupérer ou créer la feuille Excel ---
Dim w As Worksheet
Set w = getOrCreateSheet(sheetname)

' --- 3) Nettoyer la feuille et écrire le ticker en haut ---
w.Cells.Clear
w.[A1].Value = "Ticker"
w.[B1].Value = ticker

' --- 4) Déclaration des variables de travail ---
Dim r As Long
Dim c As Long
Dim key As Variant
Dim periodKey As Variant

r = 3   ' ligne de départ pour les données

' --- 5) Boucle sur chaque période (clé du JSON) ---
For Each periodKey In o.Keys
    
    ' ---- Titre principal : la période (date ou "ttm") ----
    w.Cells(r, 1).Value = "Période : " & periodKey
    w.Cells(r, 1).Font.Bold = True
    r = r + 1
    
    ' ---- Ligne des intitulés (les noms de variables) ----
    w.Cells(r, 1).Value = "Intitulés"
    w.Cells(r, 1).Font.Bold = True
    c = 2
    For Each key In o(periodKey).Keys
        w.Cells(r, c).Value = key
        c = c + 1
    Next key
    r = r + 1
    
    ' ---- Ligne des valeurs ----
    w.Cells(r, 1).Value = "Valeurs"
    w.Cells(r, 1).Font.Bold = True
    c = 2
    For Each key In o(periodKey).Keys
        w.Cells(r, c).Value = o(periodKey)(key)
        c = c + 1
    Next key
    
    ' ---- Saut de quelques lignes avant la période suivante ----
    r = r + 3

Next periodKey

' --- 6) Ajuster la largeur des colonnes ---
w.UsedRange.Columns.AutoFit



End Sub


 '*************************************
' *************************************
'ANNEXE
' *************************************
' *************************************


' *************************************
' Liste exhaustive des codes de pays
' *************************************

'GB' 'United Kingdom'
'EU' 'European Union'
'JP' 'Japan'
'DE' 'Germany'
'FR' 'France'
'CA' 'Canada'
'AU' 'Australia'
'CH' 'Switzerland'
'CN' 'China'
'NZ' 'New Zealand'
'IT' 'Italy'
'ES' 'Spain'
'SE' 'Sweden'
'NO' 'Norway'
'DK' 'Denmark'
'FI' 'Finland'
'NL' 'Netherlands'
'BE' 'Belgium'
'AT' 'Austria'
'IE' 'Ireland'
'PT' 'Portugal'
'GR' 'Greece'
'PL' 'Poland'
'CZ' 'Czech Republic'
'HU' 'Hungary'
'SK' 'Slovakia'
'SI' 'Slovenia'
'EE' 'Estonia'
'LV' 'Latvia'
'LT' 'Lithuania'
'LU' 'Luxembourg'
'MT' 'Malta'
'CY' 'Cyprus'
'BG' 'Bulgaria'
'RO' 'Romania'
'HR' 'Croatia'
'RS' 'Serbia'
'BA' 'Bosnia and Herzegovina'
'ME' 'Montenegro'
'MK' 'North Macedonia'
'AL' 'Albania'
'XK' 'Kosovo'
'MD' 'Moldova'
'UA' 'Ukraine'
'BY' 'Belarus'
'RU' 'Russia'
'TR' 'Turkey'
'IL' 'Israel'
'SA' 'Saudi Arabia'
'AE' 'United Arab Emirates'
'KW' 'Kuwait'
'QA' 'Qatar'
'BH' 'Bahrain'
'OM' 'Oman'
'JO' 'Jordan'
'LB' 'Lebanon'
'EG' 'Egypt'
'ZA' 'South Africa'
'NG' 'Nigeria'
'KE' 'Kenya'
'GH' 'Ghana'
'MA' 'Morocco'
'TN' 'Tunisia'
'DZ' 'Algeria'
'LY' 'Libya'
'SD' 'Sudan'
'ET' 'Ethiopia'
'UG' 'Uganda'
'TZ' 'Tanzania'
'RW' 'Rwanda'
'BI' 'Burundi'
'DJ' 'Djibouti'
'SO' 'Somalia'
'ER' 'Eritrea'
'SS' 'South Sudan'
'CF' 'Central African Republic'
'TD' 'Chad'
'CM' 'Cameroon'
'GA' 'Gabon'
'GQ' 'Equatorial Guinea'
'ST' 'Sao Tome and Principe'"


' *************************************
' Liste des fuseaux horaires
' *************************************

'GMT+1' 'Paris'
'GMT+2' 'AthËnes'
'GMT+3' 'Moscou'
'GMT+4' 'DubaÔ'
'GMT+5' 'Islamabad'
'GMT+6' 'Dhaka'
'GMT+7' 'Bangkok'
'GMT+8' 'PÈkin'
'GMT+9' 'Tokyo'
'GMT+10' 'Sydney'
'GMT+11' 'NoumÈa'
'GMT+12' 'Auckland'
'GMT-1' 'AÁores'
'GMT-2' 'Grenade'
'GMT-3' 'Buenos Aires'
'GMT-4' 'Santiago'
'GMT-5' 'New York'
'GMT-6' 'Mexico City'
'GMT-7' 'Denver'
'GMT-8' 'Los Angeles'
'GMT-9' 'Anchorage'
'GMT-10' 'Honolulu'
'GMT-11' 'NiuÈ'
'GMT-12' 'Baker Island'"















