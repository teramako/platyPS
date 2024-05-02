﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

using Microsoft.PowerShell.PlatyPS.MarkdownWriter;
using Microsoft.PowerShell.PlatyPS.Model;

namespace Microsoft.PowerShell.PlatyPS
{
    /// <summary>
    /// Cmdlet to generate the markdown help for commands, all commands in a module.
    /// </summary>
    [Cmdlet(VerbsCommon.New, "MarkdownCommandHelp", HelpUri = "")]
    [OutputType(typeof(FileInfo[]))]
    public sealed class NewMarkdownHelpCommand : PSCmdlet
    {
        #region cmdlet parameters

        [Parameter(Mandatory = true, ParameterSetName = "FromCommand")]
        public string[] Command { get; set; } = Array.Empty<string>();

        [Parameter()]
        [ArgumentToEncodingTransformation]
        [ArgumentEncodingCompletions]
        public System.Text.Encoding Encoding { get; set; } = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        [Parameter()]
        public SwitchParameter Force { get; set; }

        [Parameter(ParameterSetName = "FromModule")]
        public string? HelpInfoUri { get; set; }

        [Parameter(ParameterSetName = "FromModule")]
        public string? HelpVersion { get; set; }

        [Parameter(ParameterSetName = "FromModule")]
        public string? Locale { get; set; }

        [Parameter()]
        public Hashtable? Metadata { get; set; }

        [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "FromModule")]
        public string[] Module { get; set; } = Array.Empty<string>();

        [Parameter(ParameterSetName = "FromCommand")]
        public string? HelpUri { get; set; }

        [Parameter(Mandatory = true)]
        public string OutputFolder { get; set; } = Environment.CurrentDirectory;

        [Parameter(ParameterSetName = "FromModule")]
        public SwitchParameter WithModulePage { get; set; }

        [Parameter()]
        public SwitchParameter AlphabeticParamsOrder { get; set; } = true;

        [Parameter()]
        public SwitchParameter UseFullTypeName { get; set; }

        [Parameter(ParameterSetName = "FromModule")]
        [Parameter(ParameterSetName = "FromCommand")]
        public PSSession? Session { get; set; }

        [Parameter(ParameterSetName = "FromModule")]
        public string? ModulePagePath { get; set; }
        #endregion

        protected override void EndProcessing()
        {
            string fullPath = this.SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputFolder);

            if (File.Exists(fullPath))
            {
                var exception = new InvalidOperationException(string.Format(Microsoft_PowerShell_PlatyPS_Resources.PathIsNotFolder, fullPath));
                ErrorRecord err = new ErrorRecord(exception, "PathIsNotFolder", ErrorCategory.InvalidOperation, fullPath);
                ThrowTerminatingError(err);
            }

            if (!Directory.Exists(fullPath))
            {
                Directory.CreateDirectory(fullPath);
            }

            Collection<CommandHelp>? cmdHelpObjs = null;

            TransformSettings transformSettings = new TransformSettings
            {
                AlphabeticParamsOrder = AlphabeticParamsOrder,
                CreateModulePage = WithModulePage,
                DoubleDashList = false,
                ExcludeDontShow = false,
                FwLink = HelpInfoUri,
                HelpVersion = HelpVersion,
                Locale = Locale is null ? CultureInfo.GetCultureInfo("en-US") : new CultureInfo(Locale),
                ModuleGuid = null,
                ModuleName = null,
                OnlineVersionUrl = HelpUri,
                Session = Session,
                UseFullTypeName = UseFullTypeName
            };

            try
            {
                if (string.Equals(this.ParameterSetName, "FromCommand", StringComparison.OrdinalIgnoreCase))
                {
                    if (Command.Length > 0)
                    {
                        cmdHelpObjs = new TransformCommand(transformSettings).Transform(Command);
                    }
                }
                else if (string.Equals(this.ParameterSetName, "FromModule", StringComparison.OrdinalIgnoreCase))
                {
                    if (Module.Length > 0)
                    {
                        cmdHelpObjs = new TransformModule(transformSettings).Transform(Module);
                    }
                }
            }
            catch (ItemNotFoundException infe)
            {
                var exception = new ItemNotFoundException(string.Format(Microsoft_PowerShell_PlatyPS_Resources.ModuleNotFound, infe.Message));
                ErrorRecord err = new ErrorRecord(exception, "ModuleNotFound", ErrorCategory.ObjectNotFound, infe.Message);
                ThrowTerminatingError(err);
            }
            catch (CommandNotFoundException cnfe)
            {
                var exception = new CommandNotFoundException(string.Format(Microsoft_PowerShell_PlatyPS_Resources.CommandNotFound, cnfe.CommandName));
                ErrorRecord err = new ErrorRecord(exception, "CommandNotFound", ErrorCategory.ObjectNotFound, cnfe.CommandName);
                ThrowTerminatingError(err);
            }
            catch (FileNotFoundException fnfe)
            {
                var exception = new CommandNotFoundException(string.Format(Microsoft_PowerShell_PlatyPS_Resources.FileNotFound, fnfe.FileName));
                ErrorRecord err = new ErrorRecord(exception, "FileNotFound", ErrorCategory.ObjectNotFound, fnfe.FileName);
                ThrowTerminatingError(err);
            }

            if (cmdHelpObjs != null)
            {
                foreach (var cmdletHelp in cmdHelpObjs)
                {
                    var settings = new CommandHelpWriterSettings(Encoding, $"{fullPath}{Constants.DirectorySeparator}{cmdletHelp.Title}.md");
                    using var cmdWrt = new CommandHelpMarkdownWriter(settings);
                    var baseMetadata = MetadataUtils.GetCommandHelpBaseMetadata(cmdletHelp);
                    if (Metadata is null)
                    {
                        Metadata = new Hashtable(baseMetadata);
                    }
                    else
                    {
                        foreach(var metadataKey in baseMetadata.Keys)
                        {
                            if (! Metadata.ContainsKey(metadataKey))
                            {
                                Metadata.Add(metadataKey, baseMetadata[metadataKey]);
                            }
                        }
                    }

                    WriteObject(this.InvokeProvider.Item.Get(cmdWrt.Write(cmdletHelp, Metadata).FullName));
                }

                if (WithModulePage)
                {
                    string modulePagePath = ModulePagePath ?? fullPath;

                    string resolvedPathModulePagePath = this.SessionState.Path.GetUnresolvedProviderPathFromPSPath(modulePagePath);

                    var modulePageSettings = new CommandHelpWriterSettings(Encoding, resolvedPathModulePagePath);
                    using var modulePageWriter = new ModulePageWriter(modulePageSettings);

                    WriteObject(this.InvokeProvider.Item.Get(modulePageWriter.Write(cmdHelpObjs).FullName));
                }
            }
        }

        static Hashtable GetCommandHelpBaseMetadata(CommandHelp help)
        {
            var metadata = new Hashtable();
            metadata.Add("title", help.Title);
            metadata.Add("Module Name", help.ModuleName);
            metadata.Add("Locale", help.Locale.Name);
            metadata.Add("PlatyPS schema version", "2024-05-01");
            metadata.Add("HelpUri", help.OnlineVersionUrl);
            metadata.Add("ms.date", DateTime.Now.ToString("MM/dd/yyyy"));
            metadata.Add("external help file", help.ExternalHelpFile);
            return metadata;
        }

    }
}
