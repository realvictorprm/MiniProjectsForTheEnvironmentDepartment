#I @"../packages/FSharp.Data.2.3.3/lib/net40"
#r "FSharp.Data.dll"
open FSharp.Data
open System.Web.UI.WebControls

[<Literal>]
let json = 
    """
    {
      "Schutzgebiete": "Natur und Landschaft", 
      "Grundwasserstand (Meter über NN)": "Wasser", 
      "Biotope nach NatSchG und LWaldG": "Natur und Landschaft", 
      "Wasserschutzgebiete": "Wasser", 
      "FFH-Mähwiesen": "Natur und Landschaft", 
      "Grundwasserstand (m über NN)": "Wasser", 
      "FFH-Gebiete": "Natur und Landschaft", 
      "Übersicht Messreihen Luft": "Luft", 
      "Stationsvergleich": "Luft", 
      "Landschaftsschutzgebiete": "Natur und Landschaft", 
      "Komponentenvergleich": "Luft", 
      "Naturschutzgebiete": "Natur und Landschaft", 
      "Naturdenkmale, Einzelgebilde (END)": "Natur und Landschaft", 
      "Abfallbilanz absolut": "Abfall", 
      "Schutzgebietsstatistik Kreise": "Natur und Landschaft", 
      "Quellschüttung (l/s)": "Wasser", 
      "Stationsinformationen": "Luft", 
      "Überschreitungshäufigkeit": "Luft", 
      "Flächenhafte Naturdenkmale (FND)": "Natur und Landschaft", 
      "Abfallbilanz pro Kopf": "Abfall", 
      "Ortsdosisleistung - Zeitreihen": "Radioaktivität", 
      "Übersicht Messreihen Wind": "Klima und regenerative Energien", 
      "Vogelschutzgebiete (SPA)": "Natur und Landschaft", 
      "Schutzgebietsstatistik Regierungsbezirke": "Natur und Landschaft", 
      "Waldschutzgebiete": "Natur und Landschaft", 
      "Aktivitätskonzentration Radioaerosole": "Radioaktivität", 
      "Lebensraumtyp Erfassungseinheit": "Natur und Landschaft", 
      "Nationalpark": "Natur und Landschaft", 
      "Naturparks": "Natur und Landschaft"
    }"""

type OldMapping = JsonProvider<json>

let createNewMapping =
    let json = OldMapping.GetSample()
    let mapping = 
        match json.JsonValue with
        | FSharp.Data.JsonValue.Record properties ->
            [| for property in properties -> 
                match property with 
                | name, JsonValue.String topic -> name, topic 
                | _ as invalidValue-> failwithf "invalid value in mapping file: %A" invalidValue
            |]
        | _ -> failwith "invalid mapping file!"
    let grandTopics = mapping |> Array.map (fun (key, value) -> value) |> Array.distinct
    let subtopicsPerGrandTopic = 
        grandTopics |> Array.map(fun topic -> topic, mapping |> Array.filter(fun (name, t) -> t = topic) |> Array.map(fun (name, _) -> name))
    
    let generateJson =
        let block = 
            [| for (topic, mapping) in subtopicsPerGrandTopic do
                let content =
                    mapping 
                    |> Array.map(fun name -> sprintf """    "%s" """ name)
                    |> String.concat ",\n"

                yield 
                    [|  sprintf """   "%s": [ """ topic
                        content
                        "    ]"
                    |] |> String.concat "\n"
            |]
            |> String.concat ",\n" 
        
        sprintf "{\n%s\n}" block
    generateJson
    
System.IO.File.WriteAllText(@"E:/test.json", createNewMapping)

type Test = JsonProvider< @"E:/test.json">


(
    match Test.GetSample().JsonValue with
    | JsonValue.Record topics ->
        [|  for (topic, value) in topics do
                let parsedSubtopics = 
                    match value with
                    | JsonValue.Array subtopics ->
                        subtopics 
                        |> Array.map(function | JsonValue.String s -> s | _ -> failwith "invalid")

                    | _ -> failwith "invalid"
                yield topic, parsedSubtopics
        |]
    | _ -> failwith "invalid"
)