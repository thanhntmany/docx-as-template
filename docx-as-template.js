/*!
 * docx-as-template.js
 * Copyright(c) 2023 Thanhntmany
 */

'use strict';

// const process = require('process');
const fs = require('fs');
const path = require('path');


/**
 * Temporary files Cleaner
 */
const TmpDirCleaner = {
    queue: []
};

TmpDirCleaner.add = function (path) {
    this.queue.push(path);
    return this;
};

TmpDirCleaner.clean = function () {
    var curPath;
    const rm_options = {
        force: true,
        recursive: true,
    };
    while ((curPath = this.queue.shift()) !== undefined) {
        fs.rmdirSync(curPath, rm_options);
    };
};

process.on('exit', (code) => {
    TmpDirCleaner.clean();
});


/**
 * Zip handler
 */
function ZIPHandler() { };
const ZIPHandler_proto = ZIPHandler.prototype;

// #TODO:
ZIPHandler_proto.zip = function (outputPath, inputPath) { };
// #TODO:
ZIPHandler_proto.unzip = function (outputPath, inputPath) { };


/**
 * File system helper
 */
function TemplateFSHandler(basePath, backupPath) {
    this.basePath = basePath;
    this.backupPath = backupPath;

    fs.mkdirSync(this.basePath, { recursive: true });
    fs.mkdirSync(this.backupPath, { recursive: true });
};
const TemplateFSHandler_proto = TemplateFSHandler.prototype;

// #TODO:
TemplateFSHandler_proto.backup = function (path) { };
// #TODO:
TemplateFSHandler_proto.recover = function (path) { };


/**
 * Template
 */
function Template(rawFilePath, templateDir) {
    rawFilePath = path.resolve(rawFilePath);
    if (templateDir === undefined) {
        templateDir = rawFilePath + "--tmp";
        fs.mkdirSync(templateDir, { recursive: true });
        TmpDirCleaner.add(templateDir);
    };

    this.rawFilePath = rawFilePath;
    this.templateDir = templateDir;

    this.fsHandler = new TemplateFSHandler(
        path.join(this.templateDir, "base"),
        path.join(this.templateDir, "backup"),
    );

    this.zipHandler.unzip(this.fsHandler.basePath, this.rawFilePath);

    //#TODO: Clean at the end of process
};
const Template_proto = Template.prototype;

Template_proto.zipHandler = new ZIPHandler();

Template_proto.refresh = function () {
    this.fsHandler.recover();
};

// #TODO:
Template_proto.patch = function (data) {
    // #TODO:
    // this.fsHandler.backupPath( <changed-file-path> )
};

Template_proto.compile = function (outputPath) {
    return this.zipHandler.zip(outputPath, this.templateDir);
};

Template_proto.render = function (outputPath, data) {
    this.refresh();
    this.patch(data);
    return this.compile(outputPath);
};

/**
 * Core App.
 */
function App() {
    this._templateCache = {};
};
const App_proto = App.prototype;

App_proto._loadTemplate = function (templatePath) {
    console.log("_loadTemplate", templatePath);
    return new Template(templatePath);
};

App_proto.loadTemplate = function (templatePath) {
    templatePath = path.resolve(templatePath);

    return templatePath in this._templateCache
        ? this._templateCache[templatePath]
        : this._templateCache[templatePath] = this._loadTemplate(templatePath)
};

App_proto.render = function (outputPath, data, templatePath) {
    return this.loadTemplate(templatePath).render(outputPath, data);
};


/**
 * Expose `createApp()` + Core classes
 */
function createApp() {
    return new exports.App();
};

exports = module.exports = createApp;
exports.App = App;
exports.Template = Template;


/**
 * Run module as an independent application.
 */
if (require.main === module) {
    const app = createApp();

    var cmdRunner = app.cmd(args);
    cmdRunner.exec();

    //#TODO: run from param...
    console.dir(cmdRunner, { depth: null })
};
