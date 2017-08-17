let
    url = "https://stansw.visualstudio.com", 
    project = "vsts-open-in-powerbi",
    team = "vsts-open-in-powerbi Team",
    id = "70d68409-3372-43a7-9310-4dca907d5efc",

    Source = Functions[WiqlRunOneHopQueryById](url, [Project = project, Team = team], id)
in
    Source