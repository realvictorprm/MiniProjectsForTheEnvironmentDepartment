
open System
open System.IO
open System.Collections
open System.Drawing
        
let pause () = Console.ReadKey() |> ignore

open FSharp.ExcelProvider
open FSharp.Data.HttpRequestHeaders
open FSharp.Data.JsonExtensions
open FSharp.Data
open FSharp.Data.HtmlAttribute
open System.Windows.Input

type DatenQuartalsbericht = ExcelFile< @"E:\Development\Praktikum\Daten\hitliste_selektoren_udo_01_2017.xlsx", Range="A3:B100">
type MappingFile = FSharp.Data.JsonProvider< @"E:\Development\Praktikum\Daten\UDO_Topic_Mapping.json">   

type Mapping = | Mapping of Topic:string * SubTopics:string[]

type Foo = string * float

let getMapping (pathToMappingFile:string) = 
    let json = MappingFile.Load pathToMappingFile
    match json.JsonValue with
    | JsonValue.Record topics ->
        [|  for (topic, value) in topics do
                let parsedSubtopics = 
                    match value with
                    | JsonValue.Array subtopics ->
                        subtopics 
                        |> Array.map(function | JsonValue.String s -> s | _ -> failwith "invalid")

                    | _ -> failwith "invalid"
                yield Mapping(topic, parsedSubtopics)
        |]
    | _ -> failwith "invalid mapping file!"

let getSumsWithExcelFile pathToExcelFile (pathToMappingFile:string) : Foo[] =
    let file = DatenQuartalsbericht pathToExcelFile
    let printnamesValueMapping = 
        file.Data 
        |> Seq.map(fun row -> row.Printname, row.Anzahl) 
        |> Seq.filter(fun (name, _) -> name <> null)
        |> Map
    let mapping = getMapping pathToMappingFile
    
    printfn "Mapping: %A" mapping

    let sums = 
        mapping 
        |> Array.map(fun (Mapping(topic, parsedSuptopics)) -> 
            (
            topic,
            parsedSuptopics
            |> Array.sumBy(fun subtopic ->
                if printnamesValueMapping.ContainsKey subtopic then 
                    printnamesValueMapping.[subtopic] 
                else 0.
            ))
        )
    sums

open OfficeOpenXml
open OfficeOpenXml.Drawing.Chart
open System.Text.RegularExpressions
open System.Globalization
open OfficeOpenXml.Drawing
open System.Reflection

let createTable (ws:ExcelWorksheet) row column data = 
     Array.get data 0
     |> Array.iteri(fun i cell -> ws.SetValue(row, column + i, cell))
     Array.skip 1 data
     |> Array.iteri(fun i cell -> ws.SetValue(row, column + i, cell))
     ()

let monthName pathToExcelFile = 
    let regex = Regex "(?<month>\d{2})_\d{4}"
    let res = regex.Match pathToExcelFile
    if res.Groups.["month"].Success then
        match res.Groups.["month"].Value |> System.Int32.Parse with
        | value when value < 13 -> 
            DateTime(2017, value, 1).ToString("MMMM", CultureInfo.CreateSpecificCulture("de"))
        | _ -> "[MISSING]"
    else
        "[MISSING]"
[<Literal>]
let columnPrintname = 1

[<Literal>]
let columnAnzahl = 2

[<Literal>]
let columnGeneralTopicOfSubtopic = 3

[<Literal>]
let columnGeneralTopic = 5

[<Literal>]
let columnGeneralTopicCount = 6

[<EntryPoint>]
let main argv = 
    printfn "%A" argv
    printfn "Assembly location: %A" (Assembly.GetExecutingAssembly().Location)
    printfn "Current working directory: %A" (Directory.GetCurrentDirectory())
    Assembly.GetExecutingAssembly().Location |> Path.GetDirectoryName |> Directory.SetCurrentDirectory
    printfn "Corrected working directory: %A" (Directory.GetCurrentDirectory())

    let pathToExcelFile, pathToMappingFile =
        match argv with
        | [| "/h" |] -> 
            printfn "Usage: *.exe PathToExcelFile PathToTopicMappingFile or *.exe PathToExcelFile"
            exit(-1)
        | [| pathToExcelFile |] ->
            try
                pathToExcelFile |> Path.GetFullPath, Path.Combine(Directory.GetCurrentDirectory(), "mappingfile.json") |> Path.GetFullPath
            with e -> failwithf "Invalid input: %A" e
        | [| pathToExcelFile; pathToMappingFile |] ->
            try
                pathToExcelFile |> Path.GetFullPath, pathToMappingFile |> Path.GetFullPath
            with e -> failwithf "Invalid input: %A" e
        | [||] -> failwith "No input parameters specified! Type /h for help!"
        | _ -> failwithf "Unkown commandline arguments %A" argv
    
    let sumsPerTopic = getSumsWithExcelFile pathToExcelFile pathToMappingFile //@"E:\Development\Praktikum\Daten\hitliste_selektoren_udo_01_2017.xlsx" @"E:\Development\Praktikum\Daten\UDO_Topic_Mapping.json"
    
    for sum in sumsPerTopic do
        printfn "%A" sum
    
    let dateName = 
        let regex = Regex "(?<month>\d{2})_(?<year>\d{4})"
        let res = regex.Match pathToExcelFile
        if res.Groups.["month"].Success && res.Groups.["year"].Success then
            match res.Groups.["month"].Value |> System.Int32.Parse, res.Groups.["year"].Value |> System.Int32.Parse with
            | month, year when month < 13 -> 
                DateTime(year, month, 1).ToString("MMMM yyyy", CultureInfo.CreateSpecificCulture("de"))
            | _ -> "[MISSING]"
        else
            "[MISSING]"

    let testFile = sprintf "%s_generated.xlsx" (Path.GetFileNameWithoutExtension pathToExcelFile)
    File.Delete testFile
    use p = new ExcelPackage(FileInfo testFile)
    let ws = p.Workbook.Worksheets.Add "Tabelle1"
               
    let file = DatenQuartalsbericht pathToExcelFile
    let mapping = 
        getMapping pathToMappingFile 
        |> Array.collect (function | Mapping (topic, subtopics) -> subtopics |> Array.map(fun m -> m, topic))
        |> Map

    // Copy the old table and add to each sub topic it's general topic
    let filteredData = 
        file.Data
        |> Seq.filter(fun row -> row.Printname <> null)
    filteredData
    |> Seq.iteri(fun i row ->
        if row.Printname <> null then
            ws.SetValue(i + 2, columnPrintname, row.Printname)
            ws.SetValue(i + 2, columnAnzahl, row.Anzahl)
            ws.SetValue(i + 2, columnGeneralTopicOfSubtopic, mapping.[row.Printname])
    )
    ws.SetValue(1, columnPrintname, "Thema")      
    ws.SetValue(1, columnAnzahl, "Anzahl")
    ws.SetValue(1, columnGeneralTopicOfSubtopic, "Übergreifendes Thema")
    ws.Tables.Add(ExcelAddressBase(1, columnPrintname, (filteredData |> Seq.length) + 1, columnGeneralTopicOfSubtopic), "Rohdaten" ) |> ignore
    

    sumsPerTopic
    |> Array.iteri(fun i (name, sum) ->
        ws.SetValue(i + 2, columnGeneralTopic, name)
        ws.SetValue(i + 2, columnGeneralTopicCount, sum)
    )   
    ws.SetValue(1, columnGeneralTopic, "Übergreifendes Thema")      
    ws.SetValue(1, columnGeneralTopicCount, "Anzahl")
    ws.Tables.Add(ExcelAddressBase(1, columnGeneralTopic, sumsPerTopic.Length + 1, columnGeneralTopicCount), "AusgewerteteDaten" ) |> ignore
    ws.Cells.AutoFitColumns()

    let chart = ws.Drawings.AddChart("Monat " + dateName, eChartType.Pie) :?> ExcelPieChart
    chart.Title.Text <- dateName
    chart.SetSize(400, 400)
    chart.DataLabel.ShowPercent <- true
    chart.DataLabel.ShowLeaderLines <- true
    chart.Border.LineCap <- eLineCap.Flat
    chart.SetPosition(sumsPerTopic.Length + 1, 0, columnGeneralTopic - 1, 0)
    let series = 
        chart.Series.Add(                                                          
            ExcelRange.GetAddress(2, columnGeneralTopicCount, sumsPerTopic.Length + 1, columnGeneralTopicCount),
            ExcelRange.GetAddress(2, columnGeneralTopic, sumsPerTopic.Length + 1, columnGeneralTopic)
        )
    p.Save()
    
    printfn "done"
    0 // Integer-Exitcode zurückgeben
