const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map",
        entry: {
            polyfill: "@babel/polyfill",
            taskpane: "./src/taskpane/taskpane.ts",
            fallbackauthtaskpane: "./src/taskpane/fallbackauthtaskpane.ts",
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"]
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: "babel-loader"
                },
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: "ts-loader"
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader"
                },
                {
                    test: /\.(png|jpg|jpeg|gif)$/,
                    use: "file-loader"
                }
            ]
        },
        plugins: [
            new CleanWebpackPlugin(),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./src/taskpane/taskpane.html",
                chunks: ["polyfill", "taskpane"]
            }),
            new HtmlWebpackPlugin({
                filename: "fallbackauthdialog.html",
                template: "./src/taskpane/fallbackauthdialog.html",
                chunks: ["polyfill", "fallbackauthtaskpane"]
            }),
            new CopyWebpackPlugin([
                {
                    to: "taskpane.css",
                    from: "./src/taskpane/taskpane.css"
                }
            ]),
            new CopyWebpackPlugin([
                {
                    to: "assets",
                    from: "./assets"
                }
            ])
        ]
    };

    return config;
};