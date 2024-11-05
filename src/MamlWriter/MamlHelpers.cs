// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.IO;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.Linq;
using Microsoft.PowerShell.PlatyPS;
using Microsoft.PowerShell.PlatyPS.Model;
using System.Text;
using Markdig;
using Markdig.Syntax;
using System.Text.RegularExpressions;

namespace Microsoft.PowerShell.PlatyPS.MAML
{
    /// <summary>
    ///     Represents a "command:syntaxItem" element in a Powershell MAML help document.
    /// </summary>
    public class MamlConversionHelper
    {

        /// <summary>
        /// Write the help items to a file.
        /// </summary>
        public static FileInfo WriteToFile(HelpItems helpItems, string path, Encoding encoding)
        {
            var outputFile = new FileInfo(path);
            using(var writer = new StreamWriter(new FileStream(outputFile.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite), encoding))
            {
                helpItems.WriteTo(writer);
            }

            return new FileInfo(outputFile.FullName);
        }

        /// <summary>
        /// Convert a collection of platyPS CommandHelp objects into a MAML HelpItems object which can be serialized to XML.
        /// </summary>
        public static HelpItems ConvertCommandHelpToMamlHelpItems(List<CommandHelp> commandHelp)
        {
            var helpItems = new HelpItems();
            foreach(var command in commandHelp)
            {
                helpItems.Commands.Add(ConvertCommandHelpToMamlCommand(command));
            }
            return helpItems;
        }

        /// <summary>
        /// detect IndentedCode, FencedCode, Table, UnorderedList or OrderedList
        /// </summary>
        private static readonly Regex _makdownBlockReg = new(@"^(?:    |```|>|\||[\-\*]\s|[1-9]*\.)");

        /// <summary>
        /// Convert the CommandHelp object into the serializable XML
        /// </summary>
        public static Command ConvertCommandHelpToMamlCommand(CommandHelp commandHelp)
        {
            var command = new Command();
            command.Details = ConvertCommandDetails(commandHelp);
            if (commandHelp.Description is not null)
            {
                foreach (string s in Markdown.Normalize(commandHelp.Description).Split(new string[] { "\n\n" }, StringSplitOptions.None))
                {
                    // don't join lines when the block stars with IndentedCode, FencedCode, Table, UnorderedList or OrderedList
                    if (_makdownBlockReg.IsMatch(s))
                    {
                        command.Description.Add(s);
                    }
                    else
                    {
                        // XXX: Should we detect hard line break (tailing 2 spaces) ?
                        //      if should do, we will need to get rid of `TrimEnd()` from CommandHelp's parser
                        command.Description.Add(s.Replace("\n", " ").Trim());
                    }
                }
            }

            // Notes are reconstituted as Alerts in AlertSet
            // We don't break up these lines, but provide a single paragraph
            // to get better formatting.
            if (commandHelp.Notes is not null)
            {
                AlertItem alert = new();
                alert.Remark.Add(commandHelp.Notes);
                command.AlertSet.Add(alert);
            }

            if (commandHelp.Syntax is not null)
            {
                var syntaxTypeDict = new Dictionary<string, string>();
                foreach (var syntax in commandHelp.Syntax)
                {
                    foreach (var syntaxParam in syntax.SyntaxParameters)
                    {
                        if (!syntaxTypeDict.ContainsKey(syntaxParam.ParameterName))
                        {
                            syntaxTypeDict.Add(syntaxParam.ParameterName, syntaxParam.ParameterType);
                        }
                    }
                }

                foreach (var syntax in commandHelp.Syntax)
                {
                    command.Syntax.Add(ConvertSyntax(syntax, syntaxTypeDict));
                }
            }

            if (commandHelp.Examples is not null)
            {
                int exampleNumber = 0;
                foreach (var example in commandHelp.Examples)
                {
                    exampleNumber++;
                    command.Examples.Add(ConvertExample(example, exampleNumber));
                }
            }

            if (commandHelp.Parameters is not null)
            {
                foreach (var parameter in commandHelp.Parameters)
                {
                    command.Parameters.Add(ConvertParameter(parameter));
                }
            }

            if (commandHelp.Inputs is not null && commandHelp.Inputs.Count > 0)
            {
                command.InputTypes.AddRange(ConvertInputOutput(commandHelp.Inputs));
            }

            if (commandHelp.Outputs is not null && commandHelp.Outputs.Count > 0)
            {
                command.ReturnValues.AddRange(ConvertInputOutput(commandHelp.Outputs));
            }

            if (commandHelp.RelatedLinks is not null)
            {
                foreach (var link in commandHelp.RelatedLinks)
                {
                    command.RelatedLinks.Add(ConvertLink(link));
                }
            }

            return command;
        }

        private static IEnumerable<CommandValue> ConvertInputOutput(List<Model.InputOutput> inputOutput)
        {
            foreach(var io in inputOutput)
            {
                var newInputOutput = new CommandValue();
                var dataType = new DataType();
                dataType.Name = io.Typename;
                newInputOutput.DataType = dataType;
                newInputOutput.Description.Add(io.Description);
                yield return newInputOutput;
            }
        }

        private static SyntaxItem ConvertSyntax(Model.SyntaxItem syntax, IDictionary<string, string> syntaxTypeDict)
        {
            var newSyntax = new SyntaxItem();
            var firstSpace = syntax.CommandName.IndexOf(' ');
            if (firstSpace == -1)
            {
                newSyntax.CommandName = syntax.CommandName;
            }
            else
            {
                newSyntax.CommandName = syntax.CommandName.Substring(0, firstSpace);
            }
            foreach(var parameter in syntax.GetParametersInOrder())
            {
                newSyntax.Parameters.Add(ConvertParameter(parameter, syntaxTypeDict[parameter.Name]));
            }

            return newSyntax;
        }

        private static PipelineInputType GetPipelineInputType(Model.Parameter parameter)
        {

            var pipelineInput = new PipelineInputType();
            return pipelineInput;
        }

        private static ParameterValue? GetParameterValue(Model.Parameter parameter, string? syntaxParameterType = null)
        {
            // dont render <command:parameterValue> element when the parameter type is SwichParameter
            if (parameter.Type is "SwitchParameter" or "System.Management.Automation.SwitchParameter")
            {
                return null;
            }
            var parameterValue = new ParameterValue();
            if (parameter is not null)
            {
                // Set SyntaxParameter's parameter type name for <command:syntaxItem>
                parameterValue.DataType = syntaxParameterType is null ? parameter.Type : syntaxParameterType;
                if (syntaxParameterType is null)
                {
                    var t = parameter.Type;
                    parameterValue.DataType = (t.StartsWith("System.Nullable`1[", StringComparison.OrdinalIgnoreCase) && t.EndsWith("]", StringComparison.OrdinalIgnoreCase))
                        ? t.Substring(18, t.Length - 18 - 1) : t;
                }
                else
                {
                    parameterValue.DataType = syntaxParameterType;
                }
                parameterValue.IsVariableLength = parameter.VariableLength;
                parameterValue.IsMandatory = true;
            }
            return parameterValue;
        }

        private static Parameter ConvertParameter(Model.Parameter parameter, string? syntaxParameterType = null)
        {
            var newParameter = new MAML.Parameter();
            newParameter.Name = parameter.Name;
            newParameter.IsMandatory = parameter.ParameterSets.Any(x => x.IsRequired);
            newParameter.SupportsGlobbing = parameter.SupportsWildcards;
            var pSet = parameter.ParameterSets.FirstOrDefault();
            newParameter.Position = pSet is null ? Model.Constants.NamedString : pSet.Position;
            newParameter.Value = GetParameterValue(parameter, syntaxParameterType);
            newParameter.Type.Name = parameter.Type;

            if (parameter.Description is not null)
            {
                foreach(string s in parameter.Description.Split(new string[] { "\n\n" }, StringSplitOptions.None))
                {
                    newParameter.Description.Add(s.Trim());
                }
            }

            return newParameter;
        }

        private static CommandExample ConvertExample(Example example, int exampleNumber)
        {
            var newExample = new CommandExample();
            newExample.Title = string.Format($"--------- {example.Title} ---------");
            string contents = example.Remarks;
            MarkdownDocument ast = Markdig.Markdown.Parse(contents);
            foreach (Block block in ast)
            {
                if (block is CodeBlock code)
                {
                    // Aggregate multiple CodeBlocks
                    if (newExample.Code.Length > 0)
                    {
                        newExample.Code += "\n\n" + code.Lines.ToString();
                    }
                    else
                    {
                        newExample.Code = code.Lines.ToString();
                    }
                }
                else if (newExample.Code.Length == 0) // before <dev:code> => <maml:introduction>
                {
                    newExample.Description.Add(contents.Substring(block.Span.Start, block.Span.Length).Trim());
                }
                else // after <dev:code> => <dev:remarks>
                {
                    newExample.Remarks.Add(contents.Substring(block.Span.Start, block.Span.Length).Trim());
                }
            }
            if (newExample.Description.Count > 0)
            {
                // little hack: add 2 empty lines to last paragraph.
                //              otherwise, contents of <maml:introduction> and after <dev:code> are joined
                newExample.Description[newExample.Description.Count - 1] += "\n\n";
            }
            return newExample;
        }

        private static NavigationLink ConvertLink(Links link)
        {
            var newLink = new NavigationLink();
            newLink.LinkText = link.LinkText;
            newLink.Uri = link.Uri;
            return newLink;
        }

        private static CommandDetails ConvertCommandDetails(CommandHelp commandHelp)
        {
            var details = new CommandDetails();
            details.Name = commandHelp.Title;
            string[] verbNoun = commandHelp.Title.Split('-');
            details.Verb = verbNoun[0];
            details.Noun = verbNoun[1];
            details.Synopsis.Add(commandHelp.Synopsis);
            return details;
        }

    }
}

