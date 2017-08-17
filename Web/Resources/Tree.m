let
    url = "https://stansw.visualstudio.com", 
    project = "vsts-open-in-powerbi",
    team = "vsts-open-in-powerbi Team",
    id = "f015311b-b115-40a8-a848-e3fbf85d5443",

    Source = Functions[WiqlRunTreeQueryById](url, [Project = project, Team = team], id)
in
    Source