/*!
 * docx-as-template.js
 * Copyright(c) 2023 Thanhntmany
 */

'use strict';

// const process = require('process');
const fs = require('fs');
const path = require('path');
const AdmZip = require("adm-zip");

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

ZIPHandler_proto.zip = function (inputDir, outputPath) {
    // #TODO: Test to make sure work properly
    var zip = new AdmZip();
    zip.addLocalFolder(inputDir);
    zip.writeZip(outputPath);
};

ZIPHandler_proto.unzip = function (inputPath, outputDir) {
    var zip = new AdmZip(inputPath);
    zip.extractAllTo(outputDir, true);
};

/**
 * File system helper
 */
function TemplateFSHandler(baseDir, backupDir) {
    this.baseDir = baseDir;
    this.backupDir = backupDir;

    fs.mkdirSync(this.baseDir, { recursive: true });
    fs.mkdirSync(this.backupDir, { recursive: true });
};
const TemplateFSHandler_proto = TemplateFSHandler.prototype;

TemplateFSHandler_proto.getCorrespondingBackupPath = function (path) {
    var relRath = path.relative(this.baseDir, path);
    if (relRath.startsWith(".") && relRath.startsWith("..")) return path.join(this.backupDir, relRath);
};

TemplateFSHandler_proto.backupFileIfNeed = function (filePath) {
    var fileBackupPath = this.getCorrespondingBackupPath(filePath);

    // #TODO: add try catch if needed
    fs.mkdirSync(
        path.dirname(fileBackupPath),
        { recursive: true }
    );

    fs.copyFileSync(
        filePath,
        fileBackupPath,
        fs.constants.COPYFILE_EXCL
    );

};

// #TODO:
TemplateFSHandler_proto.recoverFile = function (filePath) {

};


// #TODO:
TemplateFSHandler_proto.recover = function (path) {

};


/**
 * Template
 */
function Template(rawFilePath, templateDir) {
    rawFilePath = path.resolve(rawFilePath);
    if (templateDir === undefined) {
        templateDir = rawFilePath + "--tmp";
        fs.mkdirSync(templateDir, { recursive: true });
        // #TODO: uncomment below at production.
        // TmpDirCleaner.add(templateDir);
    };

    this.rawFilePath = rawFilePath;
    this.templateDir = templateDir;

    this.fsHandler = new TemplateFSHandler(
        path.join(this.templateDir, "base"),
        path.join(this.templateDir, "backup"),
    );

    var zipFormPath = path.join(this.templateDir, "template.zip");
    fs.copyFileSync(rawFilePath, zipFormPath);

    this.zipHandler.unzip(zipFormPath, this.fsHandler.baseDir);
};
const Template_proto = Template.prototype;

Template_proto.zipHandler = new ZIPHandler();

Template_proto.refresh = function () {
    this.fsHandler.recover();
};

// #TODO:
Template_proto.patchFile = function (filePath, patcherFn) {
    this.fsHandler.backupFileIfNeed(filePath);
    // #TODO:
    // this.fsHandler.backupDir( <changed-file-path> )
};

// #TODO:
Template_proto.patch = function (data) {
    // #TODO:
    // this.fsHandler.backupDir( <changed-file-path> )
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
