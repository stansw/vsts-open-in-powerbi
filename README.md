## vsts-open-in-powerbi
Visual Studio Team Services extension which allows users to open queries in Power BI.

![Walkthrough](https://raw.githubusercontent.com/stansw/vsts-open-in-powerbi/master/Doc/walkthrough.gif)

### Acknowledgement
I would like to thank my wife [Agata](https://github.com/AgataSwierc) who had an amazing idea to create this extension and supported me in the development efforts.

### Disclaimer
This extension has been created as a personal side project and does not represent my employer's view or plans in any way.

### Development
For the short dev-loop update `vss-extension.json` and add `baseUri` option. Then you can publish extension and start dev server with `npm run dev-server`. This creates great environment for local development.

```json
"baseUri": "https://localhost:8080/",
```

```cmd
npm run dev-server
```

### Changes

2.2.0 (2018-11-15)

* Support for "dev.azure.com/{account}" URLs