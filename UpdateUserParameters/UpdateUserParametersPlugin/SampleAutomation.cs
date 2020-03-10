/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;

using System.IO.Compression;

using File = System.IO.File;
using Path = System.IO.Path;
using Directory = System.IO.Directory;

using Newtonsoft.Json;

namespace UpdateUserParametersPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public void Run(Document placeholder /*not used*/)
        {
            //LogTrace("Run called with {0}", doc.DisplayName);
            try
            {
                // !AA! Get project path and assembly from json passed in
                // !AA! Pass in output type, assembly or SVF
                using (new HeartBeat())
                {
                    string currDir = Directory.GetCurrentDirectory();

                    // Uncomment out for local debug
                    //string inputPath = System.IO.Path.Combine(currDir, @"../../inputFiles", "params.json");
                    //Dictionary<string, string> options = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText(inputPath));

                    Dictionary<string, string> options = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText("inputParams.json"));
                    string outputType = options["outputType"];
                    string inputFile = options["inputFile"];
                    string projectFile = options["projectFile"];

                    string assemblyPath = Path.GetFullPath(Path.Combine(currDir, inputFile));
                    string fullProjectPath = Path.GetFullPath(Path.Combine(currDir, projectFile));

                    if (options.ContainsKey("projectPath"))
                    {
                        // For debug of input data set
                        //DirPrint(currDir);
                        Console.WriteLine("fullProjectPath = " + fullProjectPath);

                        DesignProject dp = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
                        dp.Activate();
                    }

                    Console.WriteLine("documentPath = " + assemblyPath);
                    Document doc = inventorApplication.Documents.Open(assemblyPath);

                    // Uncomment out for local debug
                    //string paramInputPath = System.IO.Path.Combine(currDir, @"../../inputFiles", "documentParams.json");
                    //Dictionary<string, string> parameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText(paramInputPath));

                    Dictionary<string, string> parameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText("documentParams.json"));
                    foreach (KeyValuePair<string, string> entry in parameters)
                    {
                        var paramName = entry.Key;
                        var paramValue = entry.Value;
                        LogTrace($" params: {paramName}, {paramValue}");
                        ChangeParam(doc, paramName, paramValue);
                    }


                    LogTrace($"Getting full file name of assembly");
                    var docDir = Path.GetDirectoryName(doc.FullFileName);
                    var pathName = doc.FullFileName;
                    doc.Update2(true);

                    // Save Forge Viewer format (SVF)
                    string viewableDir = SaveForgeViewable(doc);
                    string viewableZip = Path.Combine(Directory.GetCurrentDirectory(), "viewable.zip");
                    ZipOutput(viewableDir, viewableZip);

                    // Save both svf and iam for now. To optimize check output type to only save one or the other
                    LogTrace($"Saving updated assembly");
                    doc.Save2(true);
                    doc.Close(true);

                    // Zip up the output assembly
                    //
                    // assembly lives in own folder under WorkingDir. Get the WorkingDir. We want to zip up the original zip to include things like project 
                    // files and libraries. For parts make sure they are also in a subdirectory
                    LogTrace("Zipping up path: " + Path.GetDirectoryName(pathName));
                    var zipInputDir = Path.GetDirectoryName(Path.GetDirectoryName(pathName) + "/../");
                    var fileName = Path.Combine(Directory.GetCurrentDirectory(), "result.zip"); // the name must be in sync with OutputIam localName in Activity
                    ZipOutput(zipInputDir, fileName);

                }
            }
            catch (Exception e)
            {
                LogError("Processing failed. " + e.ToString());
            }
        }

        public void RunWithArguments(Document placeholder, NameValueMap map)
        {
            LogTrace("RunWithArguments not implemented");
        }

        public void ChangeParam(Document doc, string paramName, string paramValue)
        {
            UserParameters userParams;

            if(doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                LogTrace("Part Document");
                PartComponentDefinition partComponentDef = ((PartDocument)doc).ComponentDefinition;
                Parameters docParams = partComponentDef.Parameters;
                userParams = docParams.UserParameters;
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
            {
                LogTrace("Assembly Document");
                AssemblyComponentDefinition assemblyComponentDef = ((AssemblyDocument)doc).ComponentDefinition;
                Parameters docParams = assemblyComponentDef.Parameters;
                userParams = docParams.UserParameters;
            }
            else
            {
                LogTrace("Unknown Document");
                // unsupported doc type, throw exception
                throw new Exception("Unsupported document type: " + doc.DocumentType.ToString());
            }

            using (new HeartBeat())
            {
                try
                {
                    LogTrace($"Setting {paramName} to {paramValue}");
                    UserParameter userParam = userParams[paramName];
                    userParam.Expression = paramValue;
                }
                catch (Exception e)
                {
                    LogError("Cannot update '{0}' parameter. ({1})", paramName, e.Message);
                }
            }
        }

        private void ZipOutput(string pathName, string fileName)
        {
            try
            {
                LogTrace($"Zipping up {fileName}");

                if (File.Exists(fileName)) File.Delete(fileName);

                // start HeartBeat around ZipFile, it could be a long operation
                using (new HeartBeat())
                {
                     ZipFile.CreateFromDirectory(pathName, fileName, CompressionLevel.Fastest, false);
                }

                LogTrace($"Saved as {fileName}");
            }
            catch (Exception e)
            {
                LogError($"********Zip Failed: {e.Message}");
            }
        }

        private string SaveForgeViewable(Document doc)
        {
            string viewableOutputDir = null;
            using (new HeartBeat())
            {
                LogTrace($"** Saving SVF");
                try
                {
                    TranslatorAddIn oAddin = null;


                    foreach (ApplicationAddIn item in inventorApplication.ApplicationAddIns)
                    {

                        if (item.ClassIdString == "{C200B99B-B7DD-4114-A5E9-6557AB5ED8EC}")
                        {
                            Trace.TraceInformation("SVF Translator addin is available");
                            oAddin = (TranslatorAddIn)item;
                            break;
                        }
                        else { }
                    }

                    if (oAddin != null)
                    {
                        Trace.TraceInformation("SVF Translator addin is available");
                        TranslationContext oContext = inventorApplication.TransientObjects.CreateTranslationContext();
                        // Setting context type
                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                        NameValueMap oOptions = inventorApplication.TransientObjects.CreateNameValueMap();
                        // Create data medium;
                        DataMedium oData = inventorApplication.TransientObjects.CreateDataMedium();

                        Trace.TraceInformation("SVF save");
                        var workingDir = Path.GetDirectoryName(doc.FullFileName);
                        var sessionDir = Path.Combine(workingDir, "SvfOutput");

                        LogTrace("*** SessionDir: " + sessionDir + ", workingDir: " + workingDir);

                        // Make sure we delete any old contents that may be in the output directory first,
                        // this is for local debugging. In DA4I the working directory is always clean
                        if (Directory.Exists(sessionDir))
                        {
                            Directory.Delete(sessionDir, true);
                        }


                        oData.FileName = Path.Combine(sessionDir, "result.collaboration");
                        var outputDir = Path.Combine(sessionDir, "output");
                        var bubbleFileOriginal = Path.Combine(outputDir, "bubble.json");
                        var bubbleFileNew = Path.Combine(sessionDir, "bubble.json");

                        // Setup SVF options
                        if (oAddin.get_HasSaveCopyAsOptions(doc, oContext, oOptions))
                        {
                            oOptions.set_Value("GeometryType", 1);
                            oOptions.set_Value("EnableExpressTranslation", true);
                            oOptions.set_Value("SVFFileOutputDir", sessionDir);
                            oOptions.set_Value("ExportFileProperties", false);
                            oOptions.set_Value("ObfuscateLabels", true);
                        }

                        LogTrace($"SVF files are oputput to: {oOptions.get_Value("SVFFileOutputDir")}");

                        oAddin.SaveCopyAs(doc, oContext, oOptions, oData);
                        Trace.TraceInformation("SVF can be exported.");
                        LogTrace($"** Saved SVF as {oData.FileName}");
                        File.Move(bubbleFileOriginal, bubbleFileNew);

                        viewableOutputDir = sessionDir;
                    }
                }
                catch (Exception e)
                {
                    LogError($"********Export to format SVF failed: {e.Message}");
                    return null;
                }
            }
            return viewableOutputDir;
        }

        static void DirPrint(string sDir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {
                        LogTrace("file: " + f);
                    }
                    DirPrint(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        #region Logging utilities

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        #endregion
    }
}