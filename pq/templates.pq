let
    Functions = 
// ----------------------------------------------------------------------------
//  Functions.VSTS.txt
// ----------------------------------------------------------------------------
let
    Referer = "vsts-open-in-powerbi/2.0.0",
    
    BatchSize = 200,
    MaxFieldsCount = 100,
    MashupTypeMap = [
        string = type text,
        integer = Int64.Type,
        dateTime = type datetime,
        plainText = type text,
        html = type text,
        treePath = type text,
        history = type text,
        double = type number,
        guid = type text,
        boolean = type logical,
        identity = type text,
        picklistString = type text,
        picklistInteger = Int64.Type,
        picklistDouble = type number
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
    ContentsWithErrorUnwrapping = (contents as function, optional options as record) as binary =>
        let
            #"Get system status codes" = { 400, 429, 500, 503, 203 },

            #"Get options" = if options <> null then options else [],
            #"Get relative path" = Record.FieldOrDefault(#"Get options", "RelativePath", ""),
            #"Get user status codes" = Record.FieldOrDefault(#"Get options", "ManualStatusHandling", {}),
            #"Get handled status codes" = List.Difference(#"Get system status codes", #"Get user status codes"),

            #"Update ManualStatusHandling" = #"Get options" & [
                ManualStatusHandling = #"Get system status codes" &  #"Get user status codes"
            ],
            #"Get contents" = contents(#"Update ManualStatusHandling"),
            #"Buffer contents" = Binary.Buffer(#"Get contents") meta Value.Metadata(#"Get contents"),
            #"Get status code" = Record.FieldOrDefault(Value.Metadata(#"Buffer contents"), "Response.Status", 0),
            #"Get error" = try Json.Document(#"Buffer contents")
                otherwise try [message = Text.FromBinary(#"Buffer contents")] 
                otherwise [],
            #"Get result" = if List.Contains(#"Get handled status codes", #"Get status code")
                    then error Error.Record("Error",
                        if #"Get status code" <> 203
                            then Record.FieldOrDefault(#"Get error", "message", "Failed to get contents from relative path'" & #"Get relative path" & "' (" & Number.ToText(#"Get status code") & ")" )
                            else "The credentials provided are invalid. Please update the credential type through a refresh or in the Data Source Settings dialog to continue.",
                    [
                        DataSourceKind = "Web.Contents", 
                        ActivityId = Diagnostics.ActivityId(),
                        RelativePath = #"Get relative path"
                    ] 
                    & Record.RemoveFields(#"Get error", {"message", "innerException", "innererror", "$id" }, MissingField.Ignore))
                    meta Value.Metadata(#"Buffer contents")
                else #"Buffer contents"
        in
            #"Get result",

    WiqlContents = (contents as function, optional options as record) as binary =>
        let
            #"Get options" = if options <> null then options else [],
            #"Get headers" = Record.FieldOrDefault(#"Get options", "Headers", []),
            #"Update headers" = #"Get options" & [
                Headers = [Referer = Referer] & #"Get headers"
            ],

            #"Get content" = ContentsWithErrorUnwrapping(contents, #"Update headers")
        in 
            #"Get content",

    WiqlQueryById = (contents as function, scope as record, id as text) =>
        let
            #"Format relative path" = FormatTeamPath(scope)
                & "/_apis/wit/wiql/" & Uri.EscapeDataString(id) 
                & "?api-version=1.0",
            #"Get result" = Json.Document(WiqlContents(contents, [RelativePath = #"Format relative path"]))
        in 
            #"Get result",

    FormatCollectionPath = (scope as record) as text => 
        if Text.Length(Text.From(Record.FieldOrDefault(scope, "Collection", ""))) > 0
            then "/" & Uri.EscapeDataString(scope[Collection])
            else "",

    FormatProjectPath = (scope as record) as text =>
        FormatCollectionPath(scope)
        & (if Text.Length(Text.From(Record.FieldOrDefault(scope, "Project", ""))) > 0 
            then "/" & Uri.EscapeDataString(scope[Project])
            else ""),

    FormatTeamPath = (scope as record) as text =>
        FormatProjectPath(scope)
        & (if Text.Length(Text.From(Record.FieldOrDefault(scope, "Project", ""))) > 0 
            and Text.Length(Text.From(Record.FieldOrDefault(scope, "Team", ""))) > 0 
            then "/" & Uri.EscapeDataString(scope[Team])
            else ""),

    GetWorkItemFieldValuesAsRecords = 
        (
            contents as function,
            scope as record,
            ids as list, 
            optional fields as list, 
            optional options as record
        ) 
        as list => 
        let
            fieldsQueryString = if (fields <> null and List.Count(fields) <= MaxFieldsCount) 
                then "&fields=" & Text.Combine(fields, ",")
                else "",
            fieldDefinitionPath = FormatCollectionPath(scope) 
                & "/_apis/wit/workitems?api-version=1.0" 
                & fieldsQueryString,
            workItemIdsBatches = Batch(ids, BatchSize),
            workItemIdsBatchQueryStrings = List.Transform(workItemIdsBatches,
                each "&ids=" & Text.Combine(List.Transform(_, Number.ToText), ",")),
            fieldValues = List.TransformMany(workItemIdsBatchQueryStrings,
                each Json.Document(WiqlContents(contents, [RelativePath = fieldDefinitionPath & _]))[value],
                (batch, _) => _)
        in
            fieldValues,

    GetWorkItemFieldValues = 
        (
            contents as function,
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

            #"Get field records" = GetWorkItemFieldValuesAsRecords(contents, scope, ids, fields, optionsEffective),
            
            #"Format field definitions path" = FormatCollectionPath(scope) & "/_apis/wit/fields?api-version=1.0",
            #"Get field definitions" = Json.Document(WiqlContents(contents, [RelativePath = #"Format field definitions path"])),
            #"Convert field definitions to table" = Table.FromRecords(#"Get field definitions"[value]),
            #"Select relevant fields" = Table.SelectRows(#"Convert field definitions to table", each List.Contains(fields, [referenceName])),
            #"Add addMashupType" = Table.AddColumn(#"Select relevant fields", "mashupType", each Record.FieldOrDefault(MashupTypeMap, [type], type any)),

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

    WiqlRunFlatWorkItemQueryById = (contents as function, url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(contents, scope, id),
            #"Get table" = WiqlRunFlatWorkItemQuery(contents, url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunFlatWorkItemQuery = (contents as function, url as text, scope as record, query as record, optional options as record) =>
        let
            #"Check queryType" = if query[queryType]? = "flat"
                then query
                else error Error.Record("Error", 
                    "Query was updated and does not return result of type ""Flat list of work items"". Please generate this file again.")
        in
            WiqlRunFlatWorkItemQueryCore(contents, url, scope, query, options),

    WiqlRunFlatWorkItemQueryCore = (contents as function, url as text, scope as record, query as record, optional options as record) =>
        let
            #"Get fields" = List.Distinct(
                { "System.Id", "System.Title", "System.WorkItemType" }
                & List.Transform(query[columns], each [referenceName])),
            #"Get ids" = List.Transform(Record.FieldOrDefault(query, "workItems", {}), each [id]),
            #"Get table" = GetWorkItemFieldValues(contents, url, scope, #"Get ids", #"Get fields", options),

            #"Add Index" = Table.AddIndexColumn(#"Get table", "Index", 1, 1),
            #"Change types" = Table.TransformColumnTypes(#"Add Index", {{"Index", Int64.Type}})
        in
            #"Change types",

    WiqlRunTreeQueryById = (contents as function, url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(contents, scope, id),
            #"Get table" = WiqlRunTreeQuery(contents, url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunTreeQuery = (contents as function, url as text, scope as record, query as record, optional options as record) =>
        let
            #"Check queryType" = if query[queryType]? = "tree" and query[queryResultType]? = "workItemLink"
                then query
                else error Error.Record("Error", 
                    "Query was updated and does not return result of type ""Tree of work items"". Please generate this file again."),

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
            #"Get table" = GetWorkItemFieldValues(contents, url, scope, #"Get ids", #"Get fields", options),

            #"Join tables" = Table.Join(#"Get table", "ID", #"Change types", "Target ID"),
            #"Remove Target ID" = Table.RemoveColumns(#"Join tables", {"Target ID"})
        in
            #"Remove Target ID",

    WiqlRunOneHopWorkItemQueryById = (contents as function, url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(contents, scope, id),
            #"Get table" = WiqlRunOneHopWorkItemQuery(contents, url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunOneHopWorkItemQuery = (contents as function, url as text, scope as record, query as record, optional options as record) as table =>
        let
            #"Check queryType" = if query[queryType]? = "oneHop" and query[queryResultType]? = "workItem"
                then query
                else error Error.Record("Error", 
                    "Query was updated and its configuration no longer matches ""Work items and direct links"""
                    & " and ""Only return items that do not have matching links""."
                    & " Please generate this file again.")
        in
            WiqlRunFlatWorkItemQueryCore(contents, url, scope, #"Check queryType", options),

    WiqlRunOneHopWorkItemLinkQueryById = (contents as function, url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(contents, scope, id),
            #"Get table" = WiqlRunOneHopWorkItemLinkQuery(contents, url, scope, #"Get query", options)
        in
            #"Get table",

    WiqlRunOneHopWorkItemLinkQuery = (contents as function, url as text, scope as record, query as record, optional options as record) as table =>
        let
            #"Check queryType" = if query[queryType]? = "oneHop" and query[queryResultType]? = "workItemLink"
                then query
                else error Error.Record("Error", 
                    "Query was updated and its configuration no longer matches ""Work items and direct links"""
                    & " and ""Return all top level items"" or ""Only return items that have matching links""."
                    & " Please generate this file again."),

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
            #"Get table" = GetWorkItemFieldValues(contents, url, scope, #"Get ids", #"Get fields", options),
            #"Join by Target ID" = Table.Join(#"Select relations", "Target ID", Table.PrefixColumns(#"Get table", "Target"), "Target.ID"),
            #"Join by Source ID" = Table.Join(#"Join by Target ID", "Source ID", Table.PrefixColumns(#"Get table", "Source"), "Source.ID"),

            #"Remove duplicate columns" = Table.RemoveColumns(#"Join by Source ID",{"Target ID", "Source ID"}),
            #"Define name conversion" = List.Transform(Table.ColumnNames(#"Remove duplicate columns"), 
                each { _, Text.Replace(Text.Replace(_, "Source.", "Source "), "Target.", "Target ")}),
            #"Apply name conversion" = Table.RenameColumns(#"Remove duplicate columns", #"Define name conversion")
        in
            #"Apply name conversion",

    WiqlRunQueryById = (contents as function, url as text, scope as record, id as text, optional options as record) as table =>
        let
            #"Get query" = WiqlQueryById(contents, scope, id),
            #"Get table" = 
                if #"Get query"[queryType]? = "flat" and #"Get query"[queryResultType]? = "workItem"
                    then WiqlRunFlatWorkItemQuery(contents, url, scope, #"Get query", options)
                else if #"Get query"[queryType]? = "oneHop" and #"Get query"[queryResultType]? = "workItemLink"
                    then WiqlRunOneHopWorkItemLinkQuery(contents, url, scope, #"Get query", options)
                else if #"Get query"[queryType]? = "oneHop" and #"Get query"[queryResultType]? = "workItem"
                    then WiqlRunOneHopWorkItemQuery(contents, url, scope, #"Get query", options)
                else if #"Get query"[queryType]? = "tree" and #"Get query"[queryResultType]? = "workItemLink"
                    then WiqlRunTreeQuery(contents, url, scope, #"Get query", options)
                 else error Error.Record("Error", 
                     "Query configuration is not supported: " 
                         & Text.From(#"Get query"[queryType]?) 
                         & " + " & Text.From(#"Get query"[queryResultType]?))
        in
            #"Get table",

    Export = [
        WiqlRunQueryById = WiqlRunQueryById,
        WiqlRunFlatWorkItemQueryById = WiqlRunFlatWorkItemQueryById,
        WiqlRunTreeWorkItemLinkQueryById = WiqlRunTreeQueryById,
        WiqlRunOneHopWorkItemQueryById = WiqlRunOneHopWorkItemQueryById,
        WiqlRunOneHopWorkItemLinkQueryById = WiqlRunOneHopWorkItemLinkQueryById
    ]
in
    Export,
// ----------------------------------------------------------------------------
    
    Query.VSTS = 
// ----------------------------------------------------------------------------
//  Query.VSTS.txt
// ----------------------------------------------------------------------------
let
    url = "https://stansw.visualstudio.com", 
    collection = "",
    project = "vsts-open-in-powerbi",
    team = "vsts-open-in-powerbi Team",
    id = "d5349265-9c9d-4808-933a-c3d27b731657",

    // Create wrapper for VSTS.Contents function as a workaround for the static code analysis in Power BI Service.
    contents = (o) => VSTS.Contents(
        url, 
        [ 
            Version              = Record.FieldOrDefault(Record.FieldOrDefault(o, "Headers", []), "Referer", ""),
            IsRetry              = Record.FieldOrDefault(o, "IsRetry", false),
            ManualStatusHandling = Record.FieldOrDefault(o, "ManualStatusHandling", {}),
            Query                = Record.FieldOrDefault(o, "Query", []),
            RelativePath         = Record.FieldOrDefault(o, "RelativePath", null),
            Timeout              = Record.FieldOrDefault(o, "Timeout", null)
       ]),

    Source = Functions[WiqlRunQueryById](contents, url, [Collection = collection, Project = project, Team = team], id)
in
    Source,
// ----------------------------------------------------------------------------

    Query.TFS = 
// ----------------------------------------------------------------------------
//  Query.TFS.txt
// ----------------------------------------------------------------------------
let
//     url = "http://localhost:8080/tfs",
//     collection = "DefaultCollection",
//     project = "proj1",
//     team = "",
//     id = "e92fecdd-fe23-4b3f-a838-1152ed54f207",
    url = "https://stansw.visualstudio.com", 
    collection = "",
    project = "vsts-open-in-powerbi",
    team = "vsts-open-in-powerbi Team",
    id = "d5349265-9c9d-4808-933a-c3d27b731657",
 
    // Create wrapper for Web.Contents function as a workaround for the static code analysis in Power BI Service.
    contents = (options) => Web.Contents(
        url, 
        [ 
            ExcludedFromCacheKey = Record.FieldOrDefault(options, "ExcludedFromCacheKey", {}),
            Headers              = Record.FieldOrDefault(options, "Headers", []),
            IsRetry              = Record.FieldOrDefault(options, "IsRetry", false),
            ManualStatusHandling = Record.FieldOrDefault(options, "ManualStatusHandling", {}),
            Query                = Record.FieldOrDefault(options, "Query", []),
            RelativePath         = Record.FieldOrDefault(options, "RelativePath", null),
            Timeout              = Record.FieldOrDefault(options, "Timeout", null)
       ]),

    Source = Functions[WiqlRunQueryById](contents, url, [Collection = collection, Project = project, Team = team], id)
in
    Source
// ----------------------------------------------------------------------------

in
    Query.TFS