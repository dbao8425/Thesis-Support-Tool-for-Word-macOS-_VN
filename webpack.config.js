/* eslint-disable no-undef */
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";

module.exports = async (env, options) => {
    const isDev = options.mode === "development";
    const config = {
        devtool: "source-map",
        entry: {
            taskpane: "./src/taskpane/taskpane.js"
        },
        output: {
            clean: true
        },
        resolve: {
            extensions: [".html", ".js"]
        },
        plugins: [
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./src/taskpane/taskpane.html",
                chunks: ["taskpane"]
            }),
            new CopyWebpackPlugin({
                patterns: [
                    {
                        from: "assets/*",
                        to: "assets/[name][ext]"
                    },
                    {
                        from: "src/taskpane/taskpane.css",
                        to: "taskpane.css"
                    },
                    {
                        from: "manifest*.xml",
                        to: "[name]" + "[ext]"
                    }
                ]
            })
        ],
        devServer: {
            hot: true,
            headers: {
                "Access-Control-Allow-Origin": "*"
            },
            server: {
                type: "https",
                options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await devCerts.getHttpsServerOptions()
            },
            port: process.env.npm_package_config_dev_server_port || 3000
        }
    };

    return config;
};
