const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const MomentLocalesPlugin = require('moment-locales-webpack-plugin');
const TerserPlugin = require('terser-webpack-plugin');
const fs = require("fs");
const webpack = require("webpack");

const urlDev="https://localhost:3000/";
const urlProd="https://sujoyu.github.io/msword-textlint/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js", ".json"]
    },
    module: {
      exprContextCritical: false,
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader", 
            options: {
              presets: [
                ["@babel/preset-env", {
                  useBuiltIns: 'entry',
                  corejs: 3,
                  targets: {
                    ie: '11',
                  },
                }]
              ]
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: "file-loader",
          options: {
            name: '[path][name].[ext]',          
          }
        }
      ]
    },
    plugins: [
      new MomentLocalesPlugin({
        localesToKeep: ['en', 'ja', 'es', 'fr', 'zh-cn', 'zh-hk', 'zh-tw',], // https://github.com/azu/textlint-rule-date-weekday-mismatch/blob/master/src/textlint-rule-date-weekday-mismatch.js#L10
      }),
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: "dict/[name].[ext]",
            from: "./node_modules/kuromoji/dict/*"
          },
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    },
    optimization: {
      minimize: true,
        minimizer: [new TerserPlugin({
          parallel: true,
          terserOptions: {
            ecma: 5,
            compress: true,
            output: {
              comments: false,
              beautify: false
            }
          }
        })]
    },
  };

  return config;
};
