var path = require("path");
var webpack = require("webpack");

module.exports = {
    target: "web",
    entry: {
        extension: "./src/extension.ts",
        app: "./src/app.ts"
    },
    output: {
        filename: "[name].bundle.js",
        libraryTarget: "amd",
        path: path.resolve(__dirname, 'out')
    },
    externals: [
        /^VSS\/.*/, 
        /^TFS\/.*/, 
        "jszip"
    ],
    resolve: {
        extensions: ["*",".webpack.js", ".web.js", ".ts", ".tsx", ".js"],
        modules: [
            path.resolve("./src"),
            path.resolve("./lib"),
            "node_modules"
        ]
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                enforce: "pre",
                loader: "tslint-loader",
                options: {
                    emitErrors: true,
                    failOnHint: true
                }
            },
            {
                test: /\.tsx?$/,
                loader: "ts-loader"
            },
            {
                test: /\.(png|jpg)$/,
                loader: 'url-loader'
            },
            {
                test: /\.css$/,
                loader: 'style-loader!css-loader'
            }
        ],
    }
}