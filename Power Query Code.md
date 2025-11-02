Query "Laufender Bezug"

  let

Quelle = List.Numbers(100, 150 ,100),

#"In Tabelle konvertiert" = Table.FromList(Quelle, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

#"Umbenannte Spalten" = Table.RenameColumns(#"In Tabelle konvertiert",{{"Column1", "Brutto"}}),

#"Hinzugefügt: Record" = Table.AddColumn(#"Umbenannte Spalten", "Record", each 
    [
    #"SV (%)"= 
            let 
            ALV = 
                
                if ([Brutto] >= 0 and [Brutto] <= 2074 ) 
                then 18.07 - 2.95

                else if ([Brutto] > 2074 and [Brutto] <= 2262 ) 
                then 18.07 - 1.95

                else if ([Brutto] > 2262  and [Brutto] <=  2451 ) 
                then 18.07 - 0.95

                else 18.07,
            GFG =
                if [Brutto]  <= 551.10 
                then 0
                else ALV
            
            in 
            GFG / 100,

    #"SV (EUR)" = 
        if [Brutto] > 6450 
        then Number.Round(6450 * #"SV (%)",2)
        else Number.Round([Brutto] * #"SV (%)",2) , 

    #"Brutto - SV" = 
        [Brutto] - #"SV (EUR)"
    ]
),

#"Erweiterte Record" = 
    Table.ExpandRecordColumn(#"Hinzugefügt: Record", "Record", 
    {"SV (%)", "SV (EUR)", "Brutto - SV"}, 
    {"SV (%)", "SV (EUR)", "Brutto - SV"}),

#"Hinzugefügt: EffektivTarifTabelle" = 
    Table.AddColumn(#"Erweiterte Record", "EffektivTarifTabelle", each

        let

            Data = tbl_EffektivTarifTabelle,
            FilteredData = Table.SelectRows( Data, (unt_tab)=> 
                [#"Brutto - SV"] >= unt_tab[#"Monatslohn VON"] and [#"Brutto - SV"] <= unt_tab[#"Monatslohn BIS"]
            )
            

        in FilteredData
),

#"Erweiterte EffektivTarifTabelle" = 
    Table.ExpandTableColumn(#"Hinzugefügt: EffektivTarifTabelle", "EffektivTarifTabelle", 
    {"Grenzsteuersatz", "Abzug", "Verkehrsabsetzbetrag", "Stufe"}, 
    {"Grenzsteuersatz", "Abzug", "Verkehrsabsetzbetrag", "Stufe"}),

#"Hinzugefügt: Lohnsteuer (EUR)" = 
    Table.AddColumn(#"Erweiterte EffektivTarifTabelle", "Lohnsteuer (EUR)", each 

    let
        LSt = ([#"Brutto - SV"] * [#"Grenzsteuersatz"]) - [#"Abzug"] - [#"Verkehrsabsetzbetrag"],
        LSt_berenigt = if LSt <0 then 0 else LSt

    in 
        Number.Round(LSt_berenigt, 2)
    , type number

),

#"Hinzugefügt: Netto" = 
    Table.AddColumn(#"Hinzugefügt: Lohnsteuer (EUR)", "Netto", each 
    Number.Round( [#"Brutto - SV"]-[#"Lohnsteuer (EUR)"] , 2)
, type number
),


#"Geänderter Typ" = 
    Table.TransformColumnTypes(#"Hinzugefügt: Netto",
    {{"Brutto", type number}, {"SV (%)", type number}, {"SV (EUR)", type number}, {"Brutto - SV", type number}, {"Grenzsteuersatz", type number}, {"Abzug", type number}, {"Verkehrsabsetzbetrag", type number}, {"Lohnsteuer (EUR)", type number}, {"Netto", type number}, {"Stufe", type text}}),

#"Sortierte Zeilen" = 
    Table.Sort(#"Geänderter Typ",
    {
        {"Brutto", Order.Descending}
        }
)

in
    #"Sortierte Zeilen"


//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Query "Sonstiger Bezug"

  let
 Quelle = #"Laufender Bezug",
    #"Sortierte Zeilen" = Table.Sort(Quelle,{{"Brutto", Order.Descending}}),
 #"Andere entfernte Spalten" = Table.SelectColumns(#"Sortierte Zeilen",{"Brutto", "SV (%)", "SV (EUR)", "Lohnsteuer (EUR)", "Netto"}),
 #"Hinzugefügt: Record" = 
    Table.AddColumn(#"Andere entfernte Spalten", "Record", each 

let 

SV_Prozent_Sonderbezug = 
if ([#"SV (%)"] -0.01) < 0 
then 0 
else ([#"SV (%)"] - 0.01)

in     
    [
    
    #"SV-BMGL 13. Bezug" = 
        if [Brutto] > 12900
        then 12900 
        else [Brutto],


    #"SV (EUR) 13. Bezug" = 
        #"SV-BMGL 13. Bezug" * SV_Prozent_Sonderbezug,

    #"Brutto - SV 13. Bezug" = 
        [Brutto] - #"SV (EUR) 13. Bezug"
    ]
    ),

#"Erweitert Record" = 
Table.ExpandRecordColumn(#"Hinzugefügt: Record", "Record", 
    {"SV-BMGL 13. Bezug", "SV (EUR) 13. Bezug", "Brutto - SV 13. Bezug"}, 
    {"SV-BMGL 13. Bezug", "SV (EUR) 13. Bezug", "Brutto - SV 13. Bezug"}
    ),

#"Hinzugefügt: LSt Sonderbezuege" = 
    Table.AddColumn( #"Erweitert Record", "LSt Sonderbezuege", each

    Table.SelectRows( tbl_LSt_Sonderbezuege, (unt_tbl)=>
    [#"Brutto - SV 13. Bezug"] >= unt_tbl[Betrag VON] and  [#"Brutto - SV 13. Bezug"] <= unt_tbl[Betrag BIS]
    )
    ),
    
#"Erweiterte LSt Sonderbezuege" = 
    Table.ExpandTableColumn(#"Hinzugefügt: LSt Sonderbezuege", "LSt Sonderbezuege", 
    {"Betrag VON", "Lohnsteuer kumuliert Vorstufe", "Lohnsteuer-Satz"}, 
    {"Betrag VON", "Lohnsteuer kumuliert Vorstufe", "Lohnsteuer-Satz"}
    ),

#"Hinzugefügt: Lohnsteuer (EUR) 13. Bezug" = 
    Table.AddColumn(#"Erweiterte LSt Sonderbezuege", "Lohnsteuer (EUR) 13. Bezug", each 
        let 
        
        Berechnung =
            [Lohnsteuer kumuliert Vorstufe] + 
            
            (
                ( [#"Brutto - SV 13. Bezug"] - [Betrag VON] ) 
                * 
                [#"Lohnsteuer-Satz"] 
            ),

        Freigrenze = 
            if [#"SV-BMGL 13. Bezug"] <= (2570 / 2)
            then 0
            else Berechnung        


        in 
        Number.Round(Freigrenze,2)
        , type number),

#"Hinzugefügt: Netto 13. Bezug" = 
    Table.AddColumn(#"Hinzugefügt: Lohnsteuer (EUR) 13. Bezug", "Netto 13. Bezug", each 
    [Brutto] - [#"SV (EUR) 13. Bezug"] - [#"Lohnsteuer (EUR) 13. Bezug"]
    , type number),
    
    
#"Hinzugefügt: SV-BMGL 14. Bezug" = 
    Table.AddColumn(#"Hinzugefügt: Netto 13. Bezug", "SV-BMGL 14. Bezug", each 
    if [#"SV-BMGL 13. Bezug"] = 12900
    then 0

    else if ( [Brutto] + [#"SV-BMGL 13. Bezug"] ) > 12900 
    then ( 12900 - [#"SV-BMGL 13. Bezug"]) 

    else if ( [Brutto] + [#"SV-BMGL 13. Bezug"] ) < 12900 
    then [#"SV-BMGL 13. Bezug"]

    else ...
    , type number),
    
    
    
    
#"Hinzugefügt: SV (EUR) 14. Bezug" = 
    Table.AddColumn(#"Hinzugefügt: SV-BMGL 14. Bezug", "SV (EUR) 14. Bezug", each 
    let 
        SV_Prozent_Sonderbezug = 
        if ([#"SV (%)"] -0.01) < 0 
        then 0 
        else ([#"SV (%)"] - 0.01)

    in 
    
        [#"SV-BMGL 14. Bezug"] * SV_Prozent_Sonderbezug
        ,type number),
    
#"Hinzugefügt: Brutto - SV 14.Bezug" = 
    Table.AddColumn(#"Hinzugefügt: SV (EUR) 14. Bezug", "Brutto - SV 14. Bezug", each 
    [Brutto] - [#"SV (EUR) 14. Bezug"], 
    type number),
    
#"Hinzugefügt: Lohnsteuerbasis Sonderbezüge Gesamt" = 
    Table.AddColumn(#"Hinzugefügt: Brutto - SV 14.Bezug", "Lohnsteuerbasis Sonderbezüge Gesamt", each 
    [#"Brutto - SV 13. Bezug"] + [#"Brutto - SV 14. Bezug"], 
    type number),
    
#"Hinzugefügt: LSt_Sonderbezuege" = 
    Table.AddColumn(#"Hinzugefügt: Lohnsteuerbasis Sonderbezüge Gesamt", "LSt_Sonderbezuege", each 
    Table.SelectRows(tbl_LSt_Sonderbezuege, (unt_tbl)=>
    [#"Lohnsteuerbasis Sonderbezüge Gesamt"] >= unt_tbl[Betrag VON] and [#"Lohnsteuerbasis Sonderbezüge Gesamt"] < unt_tbl[Betrag BIS] 
    )
),

#"Erweiterte LSt_Sonderbezuege" = 
    Table.ExpandTableColumn(#"Hinzugefügt: LSt_Sonderbezuege", "LSt_Sonderbezuege", 
    {"Betrag VON", "Lohnsteuer kumuliert Vorstufe", "Lohnsteuer-Satz"}, 
    {"Betrag VON.1", "Lohnsteuer kumuliert Vorstufe.1", "Lohnsteuer-Satz.1"}),

#"Hinzugefügt: Lohnsteuer (EUR) Sonderbezüge Gesamt" = 
    Table.AddColumn(#"Erweiterte LSt_Sonderbezuege", "Lohnsteuer (EUR) Sonderbezüge Gesamt", each 
    let

    Berechnung =
         [Lohnsteuer kumuliert Vorstufe.1] + 
            
         ( ([Lohnsteuerbasis Sonderbezüge Gesamt] - [Betrag VON.1] ) *  [#"Lohnsteuer-Satz.1"] ) ,

    Freigrenze = 
            if ( [#"SV-BMGL 13. Bezug"] + [#"SV-BMGL 14. Bezug"] ) <= 2570 
            then 0
            else Berechnung       

    in 
    Freigrenze
    , type number
),
#"Hinzugefügt: Lohnsteuer (EUR) 14. Bezug" = 
    Table.AddColumn(#"Hinzugefügt: Lohnsteuer (EUR) Sonderbezüge Gesamt", "Lohnsteuer (EUR) 14. Bezug", each 
    [#"Lohnsteuer (EUR) Sonderbezüge Gesamt"] - [#"Lohnsteuer (EUR) 13. Bezug"], 
    type number),


#"Hinzugefügt: Netto 14. Bezug" = 
    Table.AddColumn(#"Hinzugefügt: Lohnsteuer (EUR) 14. Bezug", "Netto 14. Bezug", each 
    
    let 
    Berechnung = 
    [Brutto] -
    [#"SV (EUR) 14. Bezug"] -
    [#"Lohnsteuer (EUR) 14. Bezug"]

    in 

    Number.Round(Berechnung, 2)
    ,type number),


#"Andere entfernte Spalten1" = 
    Table.SelectColumns(#"Hinzugefügt: Netto 14. Bezug",
    {"Brutto", "SV (EUR)", "Lohnsteuer (EUR)", "Netto", "SV (EUR) 13. Bezug", "Lohnsteuer (EUR) 13. Bezug", "Netto 13. Bezug", "SV (EUR) 14. Bezug", "Lohnsteuer (EUR) 14. Bezug", "Netto 14. Bezug"})


in
    #"Andere entfernte Spalten1"

//-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Query "Übersicht"
  let
    
Quelle = #"Sonstiger Bezug",

#"Hinzugefügt: Record" = 
    Table.AddColumn(Quelle, "Record", each 
    [
        
    #"Brutto 13. Bezug" = [Brutto], 
    #"Brutto 14. Bezug" = [Brutto]

    ]
),

#"Erweiterte Record" 
    = Table.ExpandRecordColumn(#"Hinzugefügt: Record", "Record", 
    {"Brutto 13. Bezug", "Brutto 14. Bezug"}, 
    {"Brutto 13. Bezug", "Brutto 14. Bezug"}
),
    

#"Hinzugefügt: Ich verdiene Brutto monatlich" = Table.AddColumn(#"Erweiterte Record", "Ich verdiene Brutto monatlich", each 
[Brutto], type number ),


#"Entpivotierte andere Spalten" = 
    Table.UnpivotOtherColumns(#"Hinzugefügt: Ich verdiene Brutto monatlich", 
    {"Ich verdiene Brutto monatlich"}, 
    "Attribut", "Wert"),


#"Hinzugefügt: Bezugsart" = 
    Table.AddColumn(#"Entpivotierte andere Spalten", "Bezugsart", each 
    if Text.Contains([Attribut], "13") 
    then "13. Bezug (Urlaubszuschuss)" 

    else if Text.Contains([Attribut], "14") 
    then "14. Bezug (Weihnachtsgeld)" 

    else "Bezug laufend"
    , type text
),


#"Hinzugefügt: Bezeichnung" = 
    Table.AddColumn(#"Hinzugefügt: Bezugsart", "Bezeichnung", each 

    if Text.Contains( [Attribut], "SV", Comparer.OrdinalIgnoreCase) 
    then "[2] Sozialversicherung"

    else if Text.Contains( [Attribut], "Lohn", Comparer.OrdinalIgnoreCase) 
    then "[3] Lohnsteuer"

    else if Text.Contains( [Attribut], "Netto", Comparer.OrdinalIgnoreCase) 
    then "[4] Netto"

    else if Text.Contains( [Attribut], "Brutto", Comparer.OrdinalIgnoreCase) 
    then "[1] Brutto"

    else [Attribut]
    , type text
),

#"Entfernte Spalten" = Table.RemoveColumns(#"Hinzugefügt: Bezeichnung",{ "Attribut"}),


#"Pivotierte Spalte" = Table.Pivot(#"Entfernte Spalten", List.Distinct(#"Entfernte Spalten"[Bezugsart]), "Bezugsart", "Wert", List.Sum),


#"Hinzugefügt: Jahresbezug" = 
    Table.AddColumn(#"Pivotierte Spalte", "Jahresbezug", each 
    12 * [Bezug laufend] +
    [#"13. Bezug (Urlaubszuschuss)"] +
    [#"14. Bezug (Weihnachtsgeld)"]
    , type number 
)

in

#"Hinzugefügt: Jahresbezug"
