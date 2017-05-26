let
    Version = "vsts-open-in-powerbi/1.0.2",
    
    BatchSize = 200,
    MaxFieldsCount = 100,
    MashupTypeMap = [
        boolean = type logical,
        dateTime = type datetime,
        double = type number,
        history = type text,
        html = type text,
        integer = Int64.Type,
        plainText = type text,
        string = type text,
        treePath = type text
    ],

    Text.FromHtml = (value as nullable text) as nullable text =>
        value,
// This section is commented out because Web.Page function is not supported in Power BI Service
// and it forces users to install Power BI Personal Gateway which is a significant effort.
//        let
//            lines = if Web.Page <> null
//                then Text.FromHtmlRec(Web.Page(value){0}[Data])
//                else { value },
//            result = if value <> null
//                then Text.Combine(lines)
//                else null
//        in
//            result,
//
//    Text.FromHtmlRec = (nodesTable as table) as list =>
//        let
//            nodes = Table.ToRecords(nodesTable),
//            lines = List.TransformMany(nodes,
//                (n) =>
//                    let
//                        text = if n[Text]? <> null then { n[Text] } else {},
//                        children = if n[Children]? <> null then @Text.FromHtmlRec(n[Children]) else {},
//                        br = if n[Name]? = "P" or n[Name]? = "BR" 
//                            then { "#(lf)" }
//                            else if n[Name]? = "DIV"
//                                then if n[Children]? <> null and Table.RowCount(n[Children]) = 1 and n[Children]{0}[Name] = "BR"
//                                    then {}
//                                    else { "#(lf)" }
//                                else {},
//                        lines = List.Combine({ text, children, br })
//                    in
//                        lines,
//                (n, t) => t)
//        in
//            lines,

    Batch = (items as list, batchSize as number) as list =>
        List.Generate(
            () => [batch = List.FirstN(items, batchSize), tail = List.Skip(items, batchSize)],
            each not List.IsEmpty([batch]),
            each [batch = List.FirstN([tail], batchSize), tail = List.Skip([tail], batchSize)],
            each [batch]),

    /// <summary>
    /// Returns the contents downloaded from a web as a binary value. In contract to standard
    /// Web.Contents this function automatically unwraps error messages for standard HTTP status
    /// codes and translates them to Power Query errors.
    /// </summary>
    ContentsWithErrorUnwrapping = (url as text, optional options as record) as binary =>
        let
            #"Get system status codes" = { 400, 429, 500, 503 },

            #"Get options" = if options <> null then options else [],
            #"Get user status codes" = Record.FieldOrDefault(#"Get options", "ManualStatusHandling", {}),
            #"Get handled status codes" = List.Difference(#"Get system status codes", #"Get user status codes"),

            #"Update ManualStatusHandling" = #"Get options" & [
                ManualStatusHandling = #"Get system status codes" &  #"Get user status codes"
            ],
            #"Get contents" = VSTS.Contents(url, #"Update ManualStatusHandling"),
            #"Buffer contents" = Binary.Buffer(#"Get contents") meta Value.Metadata(#"Get contents"),
            #"Get status code" = Record.FieldOrDefault(Value.Metadata(#"Buffer contents"), "Response.Status", 0),
            #"Get error" = try Json.Document(#"Buffer contents")
                otherwise try [message = Text.FromBinary(#"Buffer contents")] 
                otherwise [],
            #"Get result" = if List.Contains(#"Get handled status codes", #"Get status code")
                then error Error.Record("Error",
                    Record.FieldOrDefault(#"Get error", "message", "VSTS.Contents failed to get contents from '" & url & "' (" & Number.ToText(#"Get status code") & ")" ),
                    [
                        DataSourceKind = "Visual Studio Team Services", 
                        ActivityId = Diagnostics.ActivityId(),
                        Url = url
                    ] 
                    & Record.RemoveFields(#"Get error", {"message", "innerException", "innererror", "$id" }, MissingField.Ignore))
                    meta Value.Metadata(#"Buffer contents")
                else #"Buffer contents"
        in
            #"Get result",

    WiqlContents = (url as text) as binary =>
        ContentsWithErrorUnwrapping(url, [Version = Version]),

    WiqlQueryById = (url as text, scope as record, id as text) =>
        let
            #"Format url" = url
                & (if scope[Collection]? <> null then "/" & Uri.EscapeDataString(scope[Collection]) else "")
                & (if scope[Project]? <> null then "/" & Uri.EscapeDataString(scope[Project]) else "")
                & (if scope[Project]? <> null and scope[Team]? <> null then "/" & Uri.EscapeDataString(scope[Team]) else "")
                & "/_apis/wit/wiql/" & Uri.EscapeDataString(id) 
                & "?api-version=1.0",
            #"Get result" = Json.Document(WiqlContents(#"Format url"))
        in 
            #"Get result",

    GetWorkItemFieldValuesAsRecords = 
        (
            url as text,
            scope as record,
            ids as list, 
            optional fields as list, 
            optional options as record
        ) 
        as list => 
        let
            collection = Record.FieldOrDefault(scope2, "Collection", "DefaultCollection"),
            fieldsQueryString = if (fields <> null and List.Count(fields) <= MaxFieldsCount) 
                then "&fields=" & Text.Combine(fields, ",")
                else "",
            
            #"Format field definitions url" = url
                & (if scope[Collection]? <> null then "/" & Uri.EscapeDataString(scope[Collection]) else "")
                & "/_apis/wit/workitems"
                & "?api-version=1.0"
                & fieldsQueryString,
            workItemIdsBatches = Batch(ids, BatchSize),
            workItemIdsBatchQueryStrings = List.Transform(workItemIdsBatches,
                each "&ids=" & Text.Combine(List.Transform(_, Number.ToText), ",")),
            fieldValues = List.TransformMany(workItemIdsBatchQueryStrings,
                each Json.Document(WiqlContents(#"Format field definitions url" & _))[value],
                (batch, _) => _)
        in
            fieldValues,

    GetWorkItemFieldValues = 
        (
            url as text, 
            scope as record, 
            ids as list, 
            fields as list,
            optional options as record
        ) 
        as table => 
        let
            optionsDefault = [
                AsOf = null,
                ConvertHtmlToText = true,
                UseReferenceNames = false,
                AddLinkColumn = true
            ],
            optionsEffective = optionsDefault & (if options  <> null then options else []),

            #"Get field records" = GetWorkItemFieldValuesAsRecords(url, scope, ids, fields, optionsEffective),
            
            #"Format field definitions url" = url
                & (if scope[Collection]? <> null then "/" & Uri.EscapeDataString(scope[Collection]) else "")
                & "/_apis/wit/fields"
                & "?api-version=1.0",
            #"Get field definitions" = Json.Document(WiqlContents(#"Format field definitions url")),
            #"Convert field definitions to table" = Table.FromRecords(#"Get field definitions"[value]),
            #"Select relevant fields" = Table.SelectRows(#"Convert field definitions to table", each List.Contains(fields, [referenceName])),
            #"Add addMashupType" = Table.AddColumn(#"Select relevant fields", "mashupType", each Record.Field(MashupTypeMap, [type])),

            #"Convert fields to table" = Table.FromList(List.Transform(#"Get field records", each [fields]), Splitter.SplitByNothing(), { "Records" }, null, ExtraValues.Error),
            #"Expand fields records" = Table.ExpandRecordColumn(#"Convert fields to table", "Records", fields, fields),

            #"Define type conversion" = Table.ToList(Table.SelectColumns(#"Add addMashupType", {"referenceName", "mashupType"}), each {_{0}, _{1} }),
            #"Apply type conversion" = Table.TransformColumnTypes(#"Expand fields records", #"Define type conversion"),

            #"Define html conversion" = List.Transform(Table.SelectRows(#"Add addMashupType", each [type] = "html")[#"referenceName"], each { _, Text.FromHtml }),
            #"Apply html conversion" = if optionsEffective[ConvertHtmlToText]
                then Table.TransformColumns(#"Apply type conversion", #"Define html conversion")
                else #"Apply type conversion",

            #"Add Link column" = if optionsEffective[AddLinkColumn] and not Table.HasColumns(#"Apply html conversion", "Link")
                then Table.AddColumn(#"Apply html conversion", "Link", 
                    each url & "/_workitems/edit/" & Number.ToText([System.Id]), type text)
                else #"Apply html conversion",

            #"Define name conversion" = Table.ToList(Table.SelectColumns(#"Add addMashupType", {"referenceName", "name"}), each {_{0}, _{1} }),
            #"Apply name conversion" = if optionsEffective[UseReferenceNames]
                then #"Add Link column"
                else Table.RenameColumns(#"Add Link column", #"Define name conversion")
        in
            #"Apply name conversion",

    WiqlRunFlatQueryById = (url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(url, scope, id),
            #"Get table" = WiqlRunFlatQuery(url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunFlatQuery = (url as text, scope as record, query as record, optional options as record) =>
        let
            #"Check queryType" = if query[queryType]? = "flat"
                then query
                else error Error.Record("Error", 
                    "Query was updated and does not return result of type ""Flat list of work items"". Please generate this file again."),

            #"Get fields" = List.Distinct(
                { "System.Id", "System.Title", "System.WorkItemType" }
                & List.Transform(#"Check queryType"[columns], each [referenceName])),
            #"Get ids" = List.Transform(Record.FieldOrDefault(#"Check queryType", "workItems", {}), each [id]),
            #"Get table" = GetWorkItemFieldValues(url, scope, #"Get ids", #"Get fields", options),

            #"Add Index" = Table.AddIndexColumn(#"Get table", "Index", 1, 1),
            #"Change types" = Table.TransformColumnTypes(#"Add Index", {{"Index", Int64.Type}})
        in
            #"Change types",

    WiqlRunTreeQueryById = (url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(url, scope, id),
            #"Get table" = WiqlRunTreeQuery(url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunTreeQuery = (url as text, scope as record, query as record, optional options as record) =>
        let
            #"Check queryType" = if query[queryType] = "tree"
                then query
                else error Error.Record("Error", "Query was updated and does not return result of type ""Tree of work items"". Please generate this file again."),

            #"Get relations" = #"Check queryType"[workItemRelations],
            #"Convert to table" = Table.FromList(#"Get relations", Splitter.SplitByNothing(), { "Record" }, null, ExtraValues.Error),
            #"Expand relation" = Table.ExpandRecordColumn(#"Convert to table", "Record", {"target", "source"}, {"target", "source"}),
            #"Expand target" = Table.ExpandRecordColumn(#"Expand relation", "target", {"id"}, {"Target ID"}),
            #"Expand source" = Table.ExpandRecordColumn(#"Expand target", "source", {"id"}, {"Parent ID"}),
            #"Add Index" = Table.AddIndexColumn(#"Expand source", "Index", 1, 1),
            #"Change types" = Table.TransformColumnTypes(#"Add Index", {{"Index", Int64.Type}, {"Parent ID", Int64.Type}}),

            #"Get fields" = List.Distinct(
                { "System.Id", "System.Title", "System.WorkItemType" }
                & List.Transform(#"Check queryType"[columns], each [referenceName])),
            #"Get ids" = #"Change types"[Target ID],
            #"Get table" = GetWorkItemFieldValues(url, scope, #"Get ids", #"Get fields", options),

            #"Join tables" = Table.Join(#"Get table", "ID", #"Change types", "Target ID"),
            #"Remove Target ID" = Table.RemoveColumns(#"Join tables", {"Target ID"})
        in
            #"Remove Target ID",

    WiqlRunOneHopQueryById = (url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(url, scope, id),
            #"Get table" = WiqlRunOneHopQuery(url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunOneHopQuery = (url as text, scope as record, query as record, optional options as record) =>
        let
            #"Check queryType" = if query[queryType]? = "oneHop"
                then query
                else error Error.Record("Error", 
                    "Query was updated and does not return result of type ""Work items and direct links"". Please generate this file again."),

            #"Get relations" = #"Check queryType"[workItemRelations],
            #"Convert to table" = Table.FromList(#"Get relations", Splitter.SplitByNothing(), { "Record" }, null, ExtraValues.Error),
            #"Expand relation" = Table.ExpandRecordColumn(#"Convert to table", "Record", {"rel", "target", "source"}, {"Link Type", "target", "source"}),
            #"Expand target" = Table.ExpandRecordColumn(#"Expand relation", "target", {"id"}, {"Target ID"}),
            #"Expand source" = Table.ExpandRecordColumn(#"Expand target", "source", {"id"}, {"Source ID"}),
            #"Select relations" = Table.SelectRows(#"Expand source", each [Source ID] <> null and [Target ID] <> null),

            #"Get fields" = List.Distinct(
                { "System.Id", "System.Title", "System.WorkItemType" }
                & List.Transform(#"Check queryType"[columns], each [referenceName])),
            #"Get ids" = List.Distinct(#"Select relations"[Target ID] & #"Select relations"[Source ID]),
            #"Get table" = GetWorkItemFieldValues(url, scope, #"Get ids", #"Get fields", options),
            #"Join by Target ID" = Table.Join(#"Select relations", "Target ID", Table.PrefixColumns(#"Get table", "Target"), "Target.ID"),
            #"Join by Source ID" = Table.Join(#"Join by Target ID", "Source ID", Table.PrefixColumns(#"Get table", "Source"), "Source.ID"),

            #"Remove duplicate columns" = Table.RemoveColumns(#"Join by Source ID",{"Target ID", "Source ID"}),
            #"Define name conversion" = List.Transform(Table.ColumnNames(#"Remove duplicate columns"), 
                each { _, Text.Replace(Text.Replace(_, "Source.", "Source "), "Target.", "Target ")}),
            #"Apply name conversion" = Table.RenameColumns(#"Remove duplicate columns", #"Define name conversion")
        in
            #"Apply name conversion",

    WiqlRunQueryById = (url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(url, scope, id),
            #"Get table" = 
                if #"Get query"[queryType]? = "flat"
                    then WiqlRunFlatQuery(url, scope, #"Get query", options)
                else if #"Get query"[queryType]? = "oneHop"
                    then WiqlRunOneHopQuery(url, scope, #"Get query", options)
                else if #"Get query"[queryType]? = "tree"
                    then WiqlRunTreeQuery(url, scope, #"Get query", options)
                 else error Error.Record("Error", 
                    "Query type is not supported")
        in
            #"Get table",

    Export = [
        WiqlRunQueryById = WiqlRunQueryById,
        WiqlRunFlatQueryById = WiqlRunFlatQueryById,
        WiqlRunTreeQueryById = WiqlRunTreeQueryById,
        WiqlRunOneHopQueryById = WiqlRunOneHopQueryById
    ]
in
    Export