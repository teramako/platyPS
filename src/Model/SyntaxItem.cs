// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace Microsoft.PowerShell.PlatyPS.Model
{
    /// <summary>
    /// Class to represent the properties of a syntax item in PowerShell help.
    /// </summary>
    public class SyntaxItem : IEquatable<SyntaxItem>
    {
        private CommandHelp _commandHelp;

        public string CommandName { get; }
        public string ParameterSetName { get; }
        public bool HasCmdletBinding => _commandHelp.HasCmdletBinding;

        private List<SyntaxParameter> _syntaxParameters = new();
        public ReadOnlyCollection<SyntaxParameter> SyntaxParameters => _syntaxParameters.AsReadOnly();

        public SortedList<string, Parameter> Parameters => _commandHelp.Parameters;

        private HashSet<string> _parameterNames = new();

        public ReadOnlyCollection<string> ParameterNames {
            get => new ReadOnlyCollection<string>(_parameterNames.ToArray());
        }


        public bool IsDefaultParameterSet { get; }

        private bool _syntaxParametersAreSorted = true;

        public SyntaxItem(CommandHelp commandHelp, string commandName, string parameterSetName, bool isDefaultParameterSet)
        {
            _commandHelp = commandHelp;
            CommandName = commandName;
            ParameterSetName = parameterSetName;
            IsDefaultParameterSet = isDefaultParameterSet;
        }

        /// <summary>
        /// Create a copy of a syntax item.
        /// </summary>
        /// <param name="syntaxItem">The syntax item to copy.</param>
        public SyntaxItem(SyntaxItem syntaxItem)
        {
            _commandHelp = syntaxItem._commandHelp;
            CommandName = syntaxItem.CommandName;
            ParameterSetName = syntaxItem.ParameterSetName;
            IsDefaultParameterSet = syntaxItem.IsDefaultParameterSet;
            _syntaxParameters = new List<SyntaxParameter>(syntaxItem.SyntaxParameters);
            _parameterNames = new HashSet<string>(syntaxItem._parameterNames);
        }

        /// <summary>
        /// This process takes our syntax parameters and sorts them like get-command does
        /// </summary>
        public void SortParameters()
        {
            if (_syntaxParametersAreSorted)
            {
                return;
            }

            List<SyntaxParameter> sortedList = new();
            List<SyntaxParameter> positionList = new();
            List<SyntaxParameter> mandatoryList = new();
            List<SyntaxParameter> namedList = new();
            foreach(var parameter in SyntaxParameters)
            {
                if (string.Compare(parameter.Position, "named") != 0)
                {
                    positionList.Add(parameter);
                }
                else if (parameter.IsMandatory)
                {
                    mandatoryList.Add(parameter);
                }
                else
                {
                    namedList.Add(parameter);
                }
            }

            if (positionList.Count > 0)
            {
                sortedList.AddRange(positionList.OrderBy(p => int.TryParse(p.Position, out var pos) ? pos : int.MaxValue));
            }

            if (mandatoryList.Count > 0)
            {
                sortedList.AddRange(mandatoryList);
            }

            if (namedList.Count > 0)
            {
                sortedList.AddRange(namedList);
            }

            _syntaxParameters = sortedList;
            _syntaxParametersAreSorted = true;
        }

        public void AddSyntaxParameter(SyntaxParameter parameter)
        {
            string name = parameter.ParameterName;

            if (Constants.CommonParametersNames.Contains(name) && HasCmdletBinding)
            {
                return;
            }

            _parameterNames.Add(name);
            _syntaxParameters.Add(parameter);
            _syntaxParametersAreSorted = false;
        }

        private string GetFormattedSyntaxParameter(string paramName, string paramTypeName, bool isPositional, bool isRequired)
        {
            bool isSwitchParam = string.Equals(paramTypeName, "SwitchParameter", StringComparison.OrdinalIgnoreCase);
            string paramType = isSwitchParam ? string.Empty : paramTypeName;

            bool requiredPositionalSwitch = isRequired && isPositional && isSwitchParam;
            bool requiredPositional = isRequired && isPositional;
            bool requiredSwitch = isRequired && isSwitchParam;
            bool optionalSwitch = !isRequired && isSwitchParam;

            if (requiredPositionalSwitch)
            {
                return string.Format(Constants.RequiredSwitchParamTemplate, paramName, paramType);
            }
            else if (requiredPositional)
            {
                return string.Format(Constants.RequiredPositionalParamTemplate, paramName, paramType);
            }
            else if (requiredSwitch)
            {
                return string.Format(Constants.RequiredSwitchParamTemplate, paramName, paramType);
            }
            else if (isRequired)
            {
                return string.Format(Constants.RequiredParamTemplate, paramName, paramType);
            }
            else if (optionalSwitch)
            {
                return string.Format(Constants.OptionalSwitchParamTemplate, paramName, paramType);
            }
            else if (isPositional)
            {
                return string.Format(Constants.OptionalPositionalParamTemplate, paramName, paramType);
            }
            else
            {
                return string.Format(Constants.OptionalParamTemplate, paramName, paramType);
            }
        }

        public IEnumerable<Parameter> GetParametersInOrder()
        {
            SortParameters();
            foreach (var syntaxParam in SyntaxParameters)
            {
                if (Parameters.TryGetValue(syntaxParam.ParameterName, out var parameter))
                {
                    yield return parameter;
                }
            }
        }

        /// <summary>
        /// This emits the command and parameters as if they were returned by Get-Command -syntax
        /// </summary>
        /// <returns></returns>
        public string ToStringWithWrap()
        {
            SortParameters();
            StringBuilder sb = Constants.StringBuilderPool.Get();
            StringBuilder currentLine = Constants.StringBuilderPool.Get();
            currentLine.Append(CommandName);

            foreach(var parameter in SyntaxParameters)
            {
                var paramAndType = parameter.ToString();
                if (currentLine.Length + paramAndType.Length + 1 > 100)
                {
                    sb.Append(currentLine.ToString());
                    sb.AppendLine();
                    currentLine.Clear();
                }

                currentLine.Append($" {paramAndType}");
            }

            sb.Append(currentLine.ToString());

            if (HasCmdletBinding)
            {
                if (currentLine.Length + 21 > 100)
                {
                    sb.AppendLine();
                }
                sb.Append(" [<CommonParameters>]");
            }

            try
            {
                return sb.ToString();
            }
            finally
            {
                Constants.StringBuilderPool.Return(sb);
            }

        }

        public string ToSyntaxString(string fmt)
        {
            StringBuilder sb = Constants.StringBuilderPool.Get();

            try
            {
                sb.AppendFormat(fmt, ParameterSetName);
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine(Constants.CodeBlock);
                sb.AppendLine(ToStringWithWrap());
                sb.AppendLine(Constants.CodeBlock);
                return sb.ToString();
            }
            finally
            {
                Constants.StringBuilderPool.Return(sb);
            }
        }

        public override string ToString()
        {
            SortParameters();
            StringBuilder sb = Constants.StringBuilderPool.Get(); 
            try
            {
                sb.Append(CommandName);
                foreach(var parameter in SyntaxParameters)
                {
                    var paramAndType = parameter.ToString();
                    sb.Append($" {paramAndType}");
                }

                if (HasCmdletBinding)
                {
                    sb.Append($" {Constants.SyntaxCommonParameters}");
                }

                return sb.ToString();
            }
            finally
            {
                Constants.StringBuilderPool.Return(sb);
            }

        }

        public bool Equals(SyntaxItem other)
        {
            if (other is null)
            {
                return false;
            }

            return (
                string.Compare(CommandName, other.CommandName, StringComparison.CurrentCulture) == 0 &&
                string.Compare(ParameterSetName, other.ParameterSetName, StringComparison.CurrentCulture) == 0 &&
                IsDefaultParameterSet == other.IsDefaultParameterSet &&
                SyntaxParameters.Count == other.SyntaxParameters.Count &&
                SyntaxParameters.SequenceEqual<SyntaxParameter>(other.SyntaxParameters)
                );
        }

        public override bool Equals(object other)
        {
            if (other is null)
            {
                return false;
            }
            if (other is SyntaxItem syntaxItem2)
            {
                return Equals(syntaxItem2);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return (CommandName, ParameterSetName, IsDefaultParameterSet).GetHashCode();
        }

        public static bool operator ==(SyntaxItem syntaxItem1, SyntaxItem syntaxItem2)
        {
            if (syntaxItem1 is not null && syntaxItem2 is not null)
            {
                return syntaxItem1.Equals(syntaxItem2);
            }
            return false;
        }

        public static bool operator !=(SyntaxItem syntaxItem1, SyntaxItem syntaxItem2)
        {
            if (syntaxItem1 is not null && syntaxItem2 is not null)
            {
                return ! syntaxItem1.Equals(syntaxItem2);
            }
            return false;
        }
    }
}
