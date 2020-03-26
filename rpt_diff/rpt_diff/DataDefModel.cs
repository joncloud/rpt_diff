using CrystalDecisions.ReportAppServer.DataDefModel;
using ExtensionMethods;
using System;
using System.Text.Json;

namespace rpt_diff
{
    class DataDefModel
    {
        public static void ProcessCustomFunctions(CustomFunctions cfs, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("CustomFunctions");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", cfs.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (CustomFunction cf in cfs)
            {
                ProcessCustomFunction(cf, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessCustomFunction(CustomFunction cf, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Name", cf.Name);
            jsonw.WriteString("Syntax", cf.Syntax.ToStringSafe());
            jsonw.WriteString("Value", cf.Text);
            jsonw.WriteEndObject();
        }

        public static void ProcessDatabase(Database database, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Database");
            jsonw.WriteStartObject();
            ProcessTables(database.Tables, jsonw);
            ProcessTableLinks(database.TableLinks, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessTables(Tables tbls, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Tables");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", tbls.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Table tbl in tbls)
            {
                ProcessTable(tbl, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessTable(Table tbl, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Alias", tbl.Alias);
            jsonw.WriteString("Description", tbl.Description);
            jsonw.WriteString("Name", tbl.Name);
            jsonw.WriteString("QualifiedName", tbl.QualifiedName);

            ProcessFields(tbl.DataFields, jsonw, "Data");
            ProcessConnectionInfo(tbl.ConnectionInfo, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessFields(Fields flds, Utf8JsonWriter jsonw, string type)
        {
            jsonw.WritePropertyName(type+"Fields");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", flds.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Field fld in flds)
            {
                ProcessField(fld, jsonw,type);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessField(Field fld, Utf8JsonWriter jsonw,string type)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Description", fld.Description);
            jsonw.WriteString("FormulaForm", fld.FormulaForm);
            jsonw.WriteString("HeadingText", fld.HeadingText);
            jsonw.WriteString("IsRecurring", fld.IsRecurring.ToStringSafe());
            jsonw.WriteString("Kind", fld.Kind.ToStringSafe());
            jsonw.WriteString("Length", fld.Length.ToStringSafe());
            jsonw.WriteString("LongName", fld.LongName);
            jsonw.WriteString("Name", fld.Name);
            jsonw.WriteString("ShortName", fld.ShortName);
            jsonw.WriteString("Type", fld.Type.ToStringSafe());
            jsonw.WriteString("UseCount", fld.UseCount.ToStringSafe());
            switch (fld.Kind)
            {
                case CrFieldKindEnum.crFieldKindDBField:
                    {
                        DBField dbf = (DBField)fld;
                        jsonw.WriteString("TableAlias", dbf.TableAlias);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindFormulaField:
                    {
                        FormulaField ff = (FormulaField)fld;
                        jsonw.WriteString("FormulaNullTreatment", ff.FormulaNullTreatment.ToStringSafe());
                        jsonw.WriteString("Text", ff.Text.ToStringSafe());
                        jsonw.WriteString("Syntax", ff.Syntax.ToStringSafe());
                        
                        break;
                    }
                case CrFieldKindEnum.crFieldKindGroupNameField:
                    {
                        GroupNameField gnf = (GroupNameField)fld;
                        jsonw.WriteString("Group", gnf.Group.ConditionField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindParameterField:
                    {
                        ParameterField pf = (ParameterField)fld;
                        jsonw.WriteString("AllowCustomCurrentValues", pf.AllowCustomCurrentValues.ToStringSafe());
                        jsonw.WriteString("AllowMultiValue", pf.AllowMultiValue.ToStringSafe());
                        jsonw.WriteString("AllowNullValue", pf.AllowNullValue.ToStringSafe());
                        if (pf.BrowseField != null)
                        {
                            jsonw.WriteString("BrowseField", pf.BrowseField.FormulaForm);
                        }
                        jsonw.WriteString("DefaultValueSortMethod", pf.DefaultValueSortMethod.ToStringSafe());
                        jsonw.WriteString("DefaultValueSortOrder", pf.DefaultValueSortOrder.ToStringSafe());
                        jsonw.WriteString("EditMask", pf.EditMask);
                        jsonw.WriteString("IsEditableOnPanel", pf.IsEditableOnPanel.ToStringSafe());
                        jsonw.WriteString("IsOptionalPrompt", pf.IsOptionalPrompt.ToStringSafe());  
                        jsonw.WriteString("IsShownOnPanel", pf.IsShownOnPanel.ToStringSafe());

                        jsonw.WriteString("ParameterType", pf.ParameterType.ToStringSafe());
                        jsonw.WriteString("ReportName", pf.ReportName);
                        jsonw.WriteString("ValueRangeKind", pf.ValueRangeKind.ToStringSafe());


                        if (pf.CurrentValues != null)
                        {
                            ProcessValues(pf.CurrentValues, jsonw, "Current");
                        }
                        if (pf.DefaultValues != null)
                        {
                            ProcessValues(pf.DefaultValues, jsonw, "Default");
                        }
                        if (pf.InitialValues != null)
                        {
                            ProcessValues(pf.InitialValues, jsonw, "Initial");
                        }
                        if (pf.Values != null)
                        {
                            ProcessValues(pf.Values, jsonw, "");
                        }

                        jsonw.WritePropertyName("Maximum");
                        ProcessValue(pf.MaximumValue as Value, jsonw, "Maximum");
                        jsonw.WritePropertyName("Minimum");
                        ProcessValue(pf.MinimumValue as Value, jsonw, "Minimum");

                        break;   
                    }
                case CrFieldKindEnum.crFieldKindRunningTotalField:
                    {
                        RunningTotalField rtf = (RunningTotalField)fld;
                        jsonw.WriteString("EvaluateCondition", ProcessCondition(rtf.EvaluateCondition));
                        jsonw.WriteString("EvaluateConditionType", rtf.EvaluateConditionType.ToStringSafe());
                        jsonw.WriteString("Operation", rtf.Operation.ToStringSafe());
                        jsonw.WriteString("ResetCondition", ProcessCondition(rtf.ResetCondition));
                        jsonw.WriteString("ResetConditionType", rtf.ResetConditionType.ToStringSafe());
                        jsonw.WriteString("SummarizedField", rtf.SummarizedField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindSpecialField:
                    {
                        SpecialField sf = (SpecialField)fld;
                        jsonw.WriteString("SpecialType", sf.SpecialType.ToStringSafe());
                        break;
                    }
                case CrFieldKindEnum.crFieldKindSummaryField: 
                    {
                        SummaryField smf = (SummaryField)fld;
                        if (smf.Group != null)
                        {
                            jsonw.WriteString("Group", smf.Group.ConditionField.FormulaForm);
                        }
                        jsonw.WriteString("Operation", smf.Operation.ToStringSafe());
                        jsonw.WriteString("SummarizedField", smf.SummarizedField.FormulaForm);
                        break;
                    }
                case CrFieldKindEnum.crFieldKindUnknownField:
                    {
                        break;
                    }
            }
            jsonw.WriteEndObject();
        }

        private static void ProcessValues(Values values, Utf8JsonWriter jsonw, string type)
        {
            jsonw.WritePropertyName(type+"Values");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", values.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Value val in values)
            {
                ProcessValue(val, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
            
        }

        private static void ProcessValue(Value val, Utf8JsonWriter jsonw, string type="")
        {
            jsonw.WriteStartObject();
            ConstantValue cvc = val as ConstantValue;
            ExpressionValue ev = val as ExpressionValue;
            ParameterFieldDiscreteValue pfdv = val as ParameterFieldDiscreteValue;
            ParameterFieldRangeValue pfrv = val as ParameterFieldRangeValue;
            ParameterFieldValue pfv = val as ParameterFieldValue;
            if (cvc != null)
            {
                jsonw.WriteString("Value", cvc.Value.ToStringSafe());
            }
            else if (ev != null)
            {
                jsonw.WriteString("Expression", ev.Expression);
            }
            else if (pfdv != null)
            {
                jsonw.WriteString("Description", pfdv.Description);
                jsonw.WriteString("Value", Convert.ToString(pfdv.Value));
            }
            else if (pfrv != null)
            {
                jsonw.WriteString("BeginValue", Convert.ToString(pfrv.BeginValue));
                jsonw.WriteString("Description", pfrv.Description);
                jsonw.WriteString("EndValue", Convert.ToString(pfrv.EndValue));
                jsonw.WriteString("LowerBoundType", Convert.ToString(pfrv.LowerBoundType));
                jsonw.WriteString("UpperBoundType", Convert.ToString(pfrv.UpperBoundType));
            }
            else if (pfv != null)
            {
                jsonw.WriteString("Description", pfv.Description);
            }
            jsonw.WriteEndObject();
        }
        private static string ProcessCondition(object condition)
        {
            Field dfdCondition = condition as Field;
            Group gCondition = condition as Group;
            if (dfdCondition != null) // Field
            {
                return dfdCondition.FormulaForm;
            }
            else if (gCondition != null) // Group
            {
                return gCondition.ConditionField.FormulaForm;
            }
            else // Custom formula
            {
                return condition.ToStringSafe();
            }
        }

        private static void ProcessConnectionInfo(ConnectionInfo ci, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ConnectionInfo");
            jsonw.WriteStartObject();
            jsonw.WriteString("Kind", ci.Kind.ToStringSafe());
            jsonw.WriteString("Password", ci.Password);
            jsonw.WriteString("UserName", ci.UserName);
            ProcessPropertyBag(ci.Attributes, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessPropertyBag(PropertyBag pb, Utf8JsonWriter jsonw)
        {
            foreach (string pid in pb.PropertyIDs)
            {
                jsonw.WriteString(pid.Replace(" ", string.Empty), pb.StringValue[pid]);
            }
        }

        private static void ProcessTableLinks(TableLinks tls, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("TableLinks");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", tls.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (TableLink tl in tls)
            {
                ProcessTableLink(tl, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessTableLink(TableLink tl, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("JoinType", tl.JoinType.ToStringSafe());
            jsonw.WriteString("SourceTableAlias", tl.SourceTableAlias);
            jsonw.WriteString("TargetTableAlias", tl.TargetTableAlias);
            foreach (string sfn in tl.SourceFieldNames)
            {
                jsonw.WritePropertyName("SourceField");
                jsonw.WriteStartObject();
                jsonw.WriteString("Name", sfn);
                jsonw.WriteEndObject();
            }
            foreach (string tfn in tl.TargetFieldNames)
            {
                jsonw.WritePropertyName("TargetField");
                jsonw.WriteStartObject();
                jsonw.WriteString("Name", tfn);
                jsonw.WriteEndObject();
            }
            jsonw.WriteEndObject();
        }

        public static void ProcessDataDefinition(DataDefinition dd, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("DataDefinition");
            jsonw.WriteStartObject();
            ProcessFields(dd.FormulaFields, jsonw, "Formula");
            ProcessFilter(dd.GroupFilter, jsonw, "GroupFilter");
            ProcessGroups(dd.Groups, jsonw);
            ProcessFields(dd.ParameterFields, jsonw,"Parameter");
            ProcessFilter(dd.RecordFilter, jsonw, "RecordFilter");
            //ProcessFields(dd.ResultFields, jsonw, "Result"); // redundant only shows fields that are used in the report 
            ProcessFields(dd.RunningTotalFields, jsonw, "RunningTotal");
            ProcessFilter(dd.SavedDataFilter, jsonw, "SavedDataFilter");
            ProcessSorts(dd.Sorts, jsonw);
            ProcessFields(dd.SummaryFields, jsonw, "Summary");
            jsonw.WriteEndObject();
        }


        private static void ProcessFilter(Filter filter, Utf8JsonWriter jsonw,string type)
        {
            jsonw.WritePropertyName(type);
            jsonw.WriteStartObject();
            jsonw.WriteString("FormulaNullTreatment", filter.FormulaNullTreatment.ToStringSafe());
            jsonw.WriteString("FreeEditingText", filter.FreeEditingText);
            jsonw.WriteString("Name", filter.Name);
            ProcessFilterItems(filter.FilterItems, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessFilterItems(FilterItems fis, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("FilterItems");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", fis.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (FilterItem fi in fis) 
            {
                ProcessFilterItem(fi, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessFilterItem(FilterItem fi, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("ComputeText", fi.ComputeText());
            jsonw.WriteEndObject();
        }

        private static void ProcessGroups(Groups groups, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Groups");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", groups.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Group group in groups)
            {
                ProcessGroup(group, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }
        private static void ProcessGroup(Group group, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("ConditionField", group.ConditionField.FormulaForm);
            ProcessGroupOptions(group.Options, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessGroupOptions(ISCRGroupOptions go, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("GroupOptions");
            jsonw.WriteStartObject();

            jsonw.WriteString("GroupNameFormula", go.ConditionFormulas[CrGroupOptionsConditionFormulaTypeEnum.crGroupName].Text);
            jsonw.WriteString("SortDirectionFormula", go.ConditionFormulas[CrGroupOptionsConditionFormulaTypeEnum.crSortDirection].Text);
            DateGroupOptions dgo = go as DateGroupOptions;
            SpecifiedGroupOptions sgo = go as SpecifiedGroupOptions;
            if (dgo != null)
            {
                jsonw.WriteString("DateCondition", dgo.DateCondition.ToStringSafe());
            }
            if (sgo != null)
            {
                jsonw.WriteString("SpecifiedValueFilters", sgo.SpecifiedValueFilters.ToStringSafe());
                jsonw.WriteString("UnspecifiedValuesName", sgo.UnspecifiedValuesName.ToStringSafe());
                jsonw.WriteString("UnspecifiedValuesType", sgo.UnspecifiedValuesType.ToStringSafe());
            }
            jsonw.WriteEndObject();
        }

        private static void ProcessSorts(Sorts sorts, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Sorts");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", sorts.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Sort sort in sorts)
            {
                ProcessSort(sort, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessSort(Sort sort, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Direction", sort.Direction.ToStringSafe());
            jsonw.WriteString("SortField", sort.SortField.FormulaForm);
            jsonw.WriteEndObject();
        }

        public static void ProcessSummaryInfo(SummaryInfo si, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("SummaryInfo");
            jsonw.WriteStartObject();
            jsonw.WriteString("Author", si.Author);
            jsonw.WriteString("Comments", si.Comments);
            jsonw.WriteString("IsSavingWithPreview", si.IsSavingWithPreview.ToStringSafe());
            jsonw.WriteString("Keywords", si.Keywords);
            jsonw.WriteString("LastSavedBy", si.LastSavedBy);
            jsonw.WriteString("RevisionNumber", si.RevisionNumber);
            jsonw.WriteString("Subject", si.Subject);
            jsonw.WriteString("Title", si.Title);
            jsonw.WriteEndObject();
        }

    }
}
