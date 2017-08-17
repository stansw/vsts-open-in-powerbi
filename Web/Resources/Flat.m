let
    url = "https://stansw.visualstudio.com", 
    project = "vsts-open-in-powerbi",
    team = "vsts-open-in-powerbi Team",
    id = "d5349265-9c9d-4808-933a-c3d27b731657",

    Source = Functions[WiqlRunFlatQueryById](url, [Project = project, Team = team], id)
in
    Source