// This file is used by Code Analysis to maintain SuppressMessage
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given
// a specific target and scoped to a namespace, type, member, etc.

using System.Diagnostics.CodeAnalysis;

[assembly: SuppressMessage("Design", "RCS1075:Avoid empty catch clause that catches System.Exception.", Justification = "Ignore errors here", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.DeleteTempFiles")]
[assembly: SuppressMessage("Design", "RCS1075:Avoid empty catch clause that catches System.Exception.", Justification = "Ignore errors here", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.LoadExistingProteinHashFile(System.String,ValidateFastaFile.NestedStringIntList@)~System.Boolean")]
[assembly: SuppressMessage("Readability", "RCS1123:Add parentheses when necessary.", Justification = "Parentheses not needed", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.AutoFixProteinNameAndDescription(System.String@,System.String@,ValidateFastaFile.FastaValidator.ProteinNameTruncationRegex)~System.String")]
[assembly: SuppressMessage("Readability", "RCS1123:Add parentheses when necessary.", Justification = "Parentheses not needed", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.CorrectForDuplicateProteinSeqsInFasta(System.Boolean,System.Boolean,System.String,System.Collections.Generic.IList{ValidateFastaFile.ProteinHashInfo})~System.Boolean")]
[assembly: SuppressMessage("Readability", "RCS1123:Add parentheses when necessary.", Justification = "Parentheses not needed", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.EvaluateRules(System.Collections.Generic.IList{ValidateFastaFile.FastaValidator.RuleDefinitionExtended},System.String,System.String,System.Int32,System.String,System.Int32)")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_ErrorMessageTextByIndex(System.Int32,System.String)~System.String")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_ErrorsByIndex(System.Int32)~ValidateFastaFile.FastaValidator.MsgInfo")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_ErrorWarningCounts(ValidateFastaFile.FastaValidator.MsgTypeConstants,ValidateFastaFile.FastaValidator.ErrorWarningCountTypes)~System.Int32")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_FixedFASTAFileStats(ValidateFastaFile.FastaValidator.FixedFASTAFileValues)~System.Int32")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_OptionSwitch(ValidateFastaFile.FastaValidator.SwitchOptions)~System.Boolean")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_WarningMessageTextByIndex(System.Int32,System.String)~System.String")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.get_WarningsByIndex(System.Int32)~ValidateFastaFile.FastaValidator.MsgInfo")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "member", Target = "~M:ValidateFastaFile.FastaValidator.set_OptionSwitch(ValidateFastaFile.FastaValidator.SwitchOptions,System.Boolean)")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "type", Target = "~T:ValidateFastaFile.clsCustomValidateFastaFiles")]
[assembly: SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "Legacy name", Scope = "type", Target = "~T:ValidateFastaFile.clsValidateFastaFile")]
