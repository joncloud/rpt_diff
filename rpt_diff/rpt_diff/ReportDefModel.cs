
using CrystalDecisions.ReportAppServer.DataDefModel;
using CrystalDecisions.ReportAppServer.ReportDefModel;
using ExtensionMethods;
using System.Text.Json;

namespace rpt_diff
{
    class ReportDefModel
    {
        public static void ProcessReportOptions(ReportOptions ro, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ReportOptions");
            jsonw.WriteStartObject();
            jsonw.WriteString("CanSelectDistinctRecords", ro.CanSelectDistinctRecords.ToStringSafe());
            jsonw.WriteString("ConvertDateTimeType", ro.ConvertDateTimeType.ToStringSafe());
            jsonw.WriteString("ConvertNullFieldToDefault", ro.ConvertNullFieldToDefault.ToStringSafe());
            jsonw.WriteString("ConvertOtherNullsToDefault", ro.ConvertOtherNullsToDefault.ToStringSafe());
            jsonw.WriteString("EnableAsyncQuery", ro.EnableAsyncQuery.ToStringSafe());
            jsonw.WriteString("EnablePushDownGroupBy", ro.EnablePushDownGroupBy.ToStringSafe());
            jsonw.WriteString("EnableSaveDataWithReport", ro.EnableSaveDataWithReport.ToStringSafe());
            jsonw.WriteString("EnableSelectDistinctRecords", ro.EnableSelectDistinctRecords.ToStringSafe());
            jsonw.WriteString("EnableUseIndexForSpeed", ro.EnableUseIndexForSpeed.ToStringSafe());
            jsonw.WriteString("ErrorOnMaxNumOfRecords", ro.ErrorOnMaxNumOfRecords.ToStringSafe());
            jsonw.WriteString("InitialDataContext", ro.InitialDataContext);
            jsonw.WriteString("InitialReportPartName", ro.InitialReportPartName);
            jsonw.WriteString("MaxNumOfRecords", ro.MaxNumOfRecords.ToStringSafe());
            jsonw.WriteString("RefreshCEProperties", ro.RefreshCEProperties.ToStringSafe());
            jsonw.WriteEndObject();
        }

        public static void ProcessPrintOptions(PrintOptions po, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("PrintOptions");
            jsonw.WriteStartObject();
            jsonw.WriteString("DissociatePageSizeAndPrinterPaperSize", po.DissociatePageSizeAndPrinterPaperSize.ToStringSafe());
            jsonw.WriteString("DriverName", po.DriverName);
            jsonw.WriteString("PageContentHeight", po.PageContentHeight.ToStringSafe());
            jsonw.WriteString("PageContentWidth", po.PageContentWidth.ToStringSafe());
            jsonw.WriteString("PaperOrientation", po.PaperOrientation.ToStringSafe());
            jsonw.WriteString("PaperSize", po.PaperSize.ToStringSafe());
            jsonw.WriteString("PaperSource", po.PaperSource.ToStringSafe());
            jsonw.WriteString("PortName", po.PortName);
            jsonw.WriteString("PrinterDuplex", po.PrinterDuplex.ToStringSafe());
            jsonw.WriteString("PrinterName", po.PrinterName);
            jsonw.WriteString("SavedDriverName", po.SavedDriverName);
            jsonw.WriteString("SavedPaperName", po.SavedPaperName);
            jsonw.WriteString("SavedPortName", po.SavedPortName);
            jsonw.WriteString("SavedPrinterName", po.SavedPrinterName);
            ProcessPageMargins(po.PageMargins, jsonw);
            jsonw.WriteEndObject();
        }
        private static void ProcessPageMargins(PageMargins pm, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("PageMargins");
            jsonw.WriteStartObject();
            jsonw.WriteString("Bottom", pm.Bottom.ToStringSafe());
            jsonw.WriteString("Left", pm.Left.ToStringSafe());
            jsonw.WriteString("Right", pm.Right.ToStringSafe());
            jsonw.WriteString("Top", pm.Top.ToStringSafe());
            jsonw.WriteString("BottomFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeBottom].Text);
            jsonw.WriteString("LeftFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeLeft].Text);
            jsonw.WriteString("RightFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeRight].Text);
            jsonw.WriteString("TopFormula", pm.PageMarginConditionFormulas[CrPageMarginConditionFormulaTypeEnum.crPageMarginConditionFormulaTypeTop].Text);
            jsonw.WriteEndObject();
        }
        public static void ProcessSavedXMLExportFormats(XMLExportFormats xefs, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("XmlExportFormats");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", xefs.Count.ToStringSafe());
            jsonw.WriteString("DefaultExportSelection", xefs.DefaultExportSelection.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (XMLExportFormat xef in xefs)
            {
                ProcessSavedXMLExportFormat(xef, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }
        private static void ProcessSavedXMLExportFormat(XMLExportFormat xef, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Description", xef.Description);
            jsonw.WriteString("ExportBlobField", xef.ExportBlobField.ToStringSafe());
            jsonw.WriteString("FileExtension", xef.FileExtension);
            jsonw.WriteString("Name", xef.Name);
            jsonw.WriteEndObject();
        }

        public static void ProcessSubreportLinks(SubreportLinks sls, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("SubreportLinks");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", sls.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (SubreportLink sl in sls)
            {
                ProcessSubreportLink(sl, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessSubreportLink(SubreportLink sl, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("LinkedParameterName", sl.LinkedParameterName);
            jsonw.WriteString("MainReportFieldName", sl.MainReportFieldName);
            jsonw.WriteString("SubreportFieldName", sl.SubreportFieldName);
            jsonw.WriteEndObject();
        }

        public static void ProcessReportDefinition(ReportDefinition rd, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ReportDefinition");
            jsonw.WriteStartObject();
            jsonw.WriteString("ReportKind", rd.ReportKind.ToStringSafe());
            jsonw.WriteString("RptrReportFlag", rd.RptrReportFlag.ToStringSafe());
            ProcessAreas(rd.Areas, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessAreas(Areas areas, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Areas");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", areas.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Area area in areas)
            {
                ProcessArea(area, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessArea(Area area, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Kind", area.Kind.ToStringSafe());
            jsonw.WriteString("Name", area.Name);
            ProcessAreaFormat(area.Format, jsonw);
            ProcessSections(area.Sections, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessAreaFormat(ISCRAreaFormat af, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("AreaFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("EnableClampPageFooter", af.EnableClampPageFooter.ToStringSafe());
            jsonw.WriteString("EnableHideForDrillDown", af.EnableHideForDrillDown.ToStringSafe());
            jsonw.WriteString("EnableKeepTogether", af.EnableKeepTogether.ToStringSafe());
            jsonw.WriteString("EnableNewPageAfter", af.EnableNewPageAfter.ToStringSafe());
            jsonw.WriteString("EnableNewPageBefore", af.EnableNewPageBefore.ToStringSafe());
            jsonw.WriteString("EnablePrintAtBottomOfPage", af.EnablePrintAtBottomOfPage.ToStringSafe());
            jsonw.WriteString("EnableResetPageNumberAfter", af.EnableResetPageNumberAfter.ToStringSafe());
            jsonw.WriteString("EnableSuppress", af.EnableSuppress.ToStringSafe());
            jsonw.WriteString("VisibleRecordNumberPerPage", af.VisibleRecordNumberPerPage.ToStringSafe());

            jsonw.WriteString("BackgroundColor", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text);
            jsonw.WriteString("EnableClampPageFooterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text);
            jsonw.WriteString("EnableHideForDrillDownFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text);
            jsonw.WriteString("EnableKeepTogetherFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text);
            jsonw.WriteString("EnableNewPageAfterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text);
            jsonw.WriteString("EnableNewPageBeforeFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text);
            jsonw.WriteString("EnablePrintAtBottomOfPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text);
            jsonw.WriteString("EnableResetPageNumberAfterFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text);
            jsonw.WriteString("EnableSuppressFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text);
            jsonw.WriteString("EnableSuppressIfBlankFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text);
            jsonw.WriteString("EnableUnderlaySectionFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text);
            jsonw.WriteString("GroupNumberPerPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text);
            jsonw.WriteString("RecordNumberPerPageFormula", af.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessSections(Sections sections, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Sections");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", sections.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Section section in sections)
            {
                ProcessSection(section, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessSection(Section section, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Height", section.Height.ToStringSafe());
            jsonw.WriteString("Kind", section.Kind.ToStringSafe());
            jsonw.WriteString("Name", section.Name);
            ProcessReportObjects(section.ReportObjects, jsonw);
            ProcessSectionFormat(section.Format, jsonw);
            jsonw.WriteEndObject();
        }
        private static void ProcessReportObjects(ReportObjects ros, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ReportObjects");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", ros.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (ReportObject ro in ros)
            {
                ProcessReportObject(ro, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }
        private static void ProcessReportObject(ReportObject ro, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            // for all types the same properties

            jsonw.WriteString("Height", ro.Height.ToStringSafe());
            jsonw.WriteString("Kind", ro.Kind.ToStringSafe());
            jsonw.WriteString("Left", ro.Left.ToStringSafe());
            jsonw.WriteString("LinkedURI", ro.LinkedURI);
            jsonw.WriteString("Name", ro.Name);
            //jsonw.WriteString("ReportPartIDDataContext", ro.ReportPartBookmark.ReportPartID.DataContext);
            //jsonw.WriteString("ReportPartIDName", ro.ReportPartBookmark.ReportPartID.Name);
            //jsonw.WriteString("ReportURI", ro.ReportPartBookmark.ReportURI);
            jsonw.WriteString("Top", ro.Top.ToStringSafe());
            jsonw.WriteString("Width", ro.Width.ToStringSafe());
            // kind specific properties
            switch (ro.Kind)
            {
                case CrReportObjectKindEnum.crReportObjectKindBlobField:
                    {
                        BlobFieldObject bfo = (BlobFieldObject)ro;
                        jsonw.WriteString("DataSourceName", bfo.DataSourceName);
                        jsonw.WriteString("OriginalHeight", bfo.OriginalHeight.ToStringSafe());
                        jsonw.WriteString("OriginalWidth", bfo.OriginalWidth.ToStringSafe());
                        jsonw.WriteString("XScaling", bfo.XScaling.ToStringSafe());
                        jsonw.WriteString("YScaling", bfo.YScaling.ToStringSafe());
                        ProcessPictureFormat(bfo.PictureFormat,jsonw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindBox: 
                    {
                        BoxObject bo = (BoxObject)ro;
                        jsonw.WriteString("Bottom", bo.Bottom.ToStringSafe());
                        jsonw.WriteString("CornerEllipseHeight", bo.CornerEllipseHeight.ToStringSafe());
                        jsonw.WriteString("CornerEllipseWidth", bo.CornerEllipseWidth.ToStringSafe());
                        jsonw.WriteString("EnableExtendToBottomOfSection", bo.EnableExtendToBottomOfSection.ToStringSafe());
                        jsonw.WriteString("EndSectionName", bo.EndSectionName);
                        jsonw.WriteString("FillColor", bo.FillColor.ToStringSafe());
                        jsonw.WriteString("LineColor", bo.LineColor.ToStringSafe());
                        jsonw.WriteString("LineStyle", bo.LineStyle.ToStringSafe());
                        jsonw.WriteString("LineThickness", bo.LineThickness.ToStringSafe());
                        jsonw.WriteString("Right", bo.Right.ToStringSafe());
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindFieldHeading:
                    {
                        FieldHeadingObject fho = (FieldHeadingObject)ro;
                        jsonw.WriteString("FieldObjectName", fho.FieldObjectName);
                        jsonw.WriteString("MaxNumberOfLines", fho.MaxNumberOfLines.ToStringSafe());
                        jsonw.WriteString("ReadingOrder", fho.ReadingOrder.ToStringSafe());
                        jsonw.WriteString("Text", fho.Text);
                        ProcessFontColor(fho.FontColor, jsonw);
                        ProcessParagraphs(fho.Paragraphs, jsonw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindField:
                    {
                        FieldObject fo = (FieldObject)ro;
                        jsonw.WriteString("DataSourceName", fo.DataSourceName);
                        jsonw.WriteString("FieldValueType", fo.FieldValueType.ToStringSafe());
                        ProcessFieldFormat(fo.FieldFormat, jsonw, fo.FieldValueType);
                        ProcessFontColor(fo.FontColor, jsonw);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindLine:
                    {
                        LineObject lo = (LineObject)ro;
                        jsonw.WriteString("Bottom", lo.Bottom.ToStringSafe());
                        jsonw.WriteString("EnableExtendToBottomOfSection", lo.EnableExtendToBottomOfSection.ToStringSafe());
                        jsonw.WriteString("EndSectionName", lo.EndSectionName);
                        jsonw.WriteString("LineColor", lo.LineColor.ToStringSafe());
                        jsonw.WriteString("LineStyle", lo.LineStyle.ToStringSafe());
                        jsonw.WriteString("LineThickness", lo.LineThickness.ToStringSafe());
                        jsonw.WriteString("Right", lo.Right.ToStringSafe());
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindSubreport:
                    {
                        SubreportObject so = (SubreportObject)ro;
                        jsonw.WriteString("EnableOnDemand", so.EnableOnDemand.ToStringSafe());
                        jsonw.WriteString("IsImported", so.IsImported.ToStringSafe());
                        jsonw.WriteString("SubreportLocation", so.SubreportLocation);
                        jsonw.WriteString("SubreportName", so.SubreportName);
                        break;
                    }
                case CrReportObjectKindEnum.crReportObjectKindText:
                    {
                        TextObject to = (TextObject)ro;
                        jsonw.WriteString("MaxNumberOfLines", to.MaxNumberOfLines.ToStringSafe());
                        jsonw.WriteString("ReadingOrder", to.ReadingOrder.ToStringSafe());
                        jsonw.WriteString("Text", to.Text);
                        ProcessFontColor(to.FontColor, jsonw);
                        ProcessParagraphs(to.Paragraphs, jsonw);
                        break;
                    }
                default:
                    {
                        break;
                    }
            }
            ProcessBorder(ro.Border, jsonw);
            ProcessObjectFormat(ro.Format, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessParagraphs(Paragraphs paragraphs, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Paragraphs");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", paragraphs.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (Paragraph p in paragraphs)
            {
                ProcessParagraph(p, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessParagraph(Paragraph p, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Alignment", p.Alignment.ToStringSafe());
            jsonw.WriteString("ReadingOrder", p.ReadingOrder.ToStringSafe());
            ProcessFontColor(p.FontColor, jsonw);
            ProcessIndentAndSpacingFormat(p.IndentAndSpacingFormat, jsonw);
            ProcessParagraphElements(p.ParagraphElements, jsonw);
            ProcessTabStops(p.TabStops, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessParagraphElements(ParagraphElements pe, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ParagraphElements");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", pe.Count.ToStringSafe());
            jsonw.WritePropertyName("Items");
            jsonw.WriteStartArray();
            foreach (ParagraphElement p in pe)
            {
                ProcessParagraphElement(p, jsonw);
            }
            jsonw.WriteEndArray();
            jsonw.WriteEndObject();
        }

        private static void ProcessParagraphElement(ParagraphElement p, Utf8JsonWriter jsonw)
        {
            jsonw.WriteStartObject();
            jsonw.WriteString("Kind", p.Kind.ToStringSafe());
            ProcessFontColor(p.FontColor, jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessFontColor(FontColor fc, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("FontColor");
            jsonw.WriteStartObject();
            jsonw.WriteString("Color", fc.Color.ToStringSafe());
            jsonw.WriteString("Bold", fc.Font.Bold.ToStringSafe());
            jsonw.WriteString("Charset", fc.Font.Charset.ToStringSafe());
            jsonw.WriteString("Italic", fc.Font.Italic.ToStringSafe());
            jsonw.WriteString("Name", fc.Font.Name);
            jsonw.WriteString("Size", fc.Font.Size.ToStringSafe());
            jsonw.WriteString("Strikethrough", fc.Font.Strikethrough.ToStringSafe());
            jsonw.WriteString("Underline", fc.Font.Underline.ToStringSafe());
            jsonw.WriteString("Weight", fc.Font.Weight.ToStringSafe());
            jsonw.WriteString("ColorFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeColor].Text);
            jsonw.WriteString("NameFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeName].Text);
            jsonw.WriteString("SizeFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeSize].Text);
            jsonw.WriteString("StrikeoutFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStrikeout].Text);
            jsonw.WriteString("StyleFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStyle].Text);
            jsonw.WriteString("UnderlineFormula", fc.ConditionFormulas[CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeUnderline].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessTabStops(TabStops ts, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("TabStops");
            jsonw.WriteStartObject();
            jsonw.WriteString("Count", ts.Count.ToStringSafe());
            foreach (TabStop t in ts)
            {
                ProcessTabStop(t, jsonw);
            }
            jsonw.WriteEndObject();
        }

        private static void ProcessTabStop(TabStop t, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("TabStop");
            jsonw.WriteStartObject();
            jsonw.WriteString("Alignment", t.Alignment.ToStringSafe());
            jsonw.WriteString("XOffset", t.XOffset.ToStringSafe());
            jsonw.WriteEndObject();
        }

        private static void ProcessPictureFormat(PictureFormat pf, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("PictureFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("BottomCropping", pf.BottomCropping.ToStringSafe());
            jsonw.WriteString("LeftCropping", pf.LeftCropping.ToStringSafe());
            jsonw.WriteString("RightCropping", pf.RightCropping.ToStringSafe());
            jsonw.WriteString("TopCropping", pf.TopCropping.ToStringSafe());
            jsonw.WriteEndObject();
        }

        private static void ProcessFieldFormat(FieldFormat ff, Utf8JsonWriter jsonw, CrFieldValueTypeEnum vt)
        {
            switch (vt)
            {
                case CrFieldValueTypeEnum.crFieldValueTypeBooleanField:
                    {
                        ProcessBooleanFormat(ff.BooleanFormat, jsonw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeDateField:
                    {
                        ProcessDateFormat(ff.DateFormat, jsonw); 
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeDateTimeField:
                    {
                        ProcessDateTimeFormat(ff.DateTimeFormat, jsonw);    
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeCurrencyField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt16sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt16uField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt32sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt32uField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt8sField:
                case CrFieldValueTypeEnum.crFieldValueTypeInt8uField:
                case CrFieldValueTypeEnum.crFieldValueTypeNumberField:
                    {
                        ProcessNumericFormat(ff.NumericFormat, jsonw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeStringField:
                    {
                        ProcessStringFormat(ff.StringFormat, jsonw);
                        break;
                    }
                case CrFieldValueTypeEnum.crFieldValueTypeTimeField:
                    {
                        ProcessTimeFormat(ff.TimeFormat, jsonw);
                        break;
                    }
                default:
                    {
                        
                        break;
                    }
            }
            ProcessCommonFormat(ff.CommonFormat, jsonw);
        }

        private static void ProcessBooleanFormat(BooleanFieldFormat bff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("BooleanFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("OutputFormat", bff.OutputFormat.ToStringSafe());
            jsonw.WriteString("OutputFormatFormula", bff.ConditionFormulas[CrBooleanFieldFormatConditionFormulaTypeEnum.crBooleanFieldFormatConditionFormulaTypeOutputFormat].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessCommonFormat(CommonFieldFormat cff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("CommonFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("EnableSuppressIfDuplicated", cff.EnableSuppressIfDuplicated.ToStringSafe());
            jsonw.WriteString("EnableSystemDefault", cff.EnableSystemDefault.ToStringSafe());
            jsonw.WriteString("EnableSuppressIfDuplicatedFormula", cff.ConditionFormulas[CrCommonFieldFormatConditionFormulaTypeEnum.crCommonFieldFormatConditionFormulaTypeSuppressIfDuplicated].Text);
            jsonw.WriteString("EnableSystemDefaultFormula", cff.ConditionFormulas[CrCommonFieldFormatConditionFormulaTypeEnum.crCommonFieldFormatConditionFormulaTypeUseSystemDefault].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessDateTimeFormat(DateTimeFieldFormat dtff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("DateTimeFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("DateTimeSeparator", dtff.DateTimeSeparator);
            jsonw.WriteString("DateTimeOrder", dtff.DateTimeOrder.ToStringSafe());
            jsonw.WriteString("DateTimeSeparatorFormula", dtff.ConditionFormulas[CrDateTimeFieldFormatConditionFormulaTypeEnum.crDateTimeFieldFormatConditionFormulaTypeDateTimeOrder].Text);
            jsonw.WriteString("DateTimeOrderFormula", dtff.ConditionFormulas[CrDateTimeFieldFormatConditionFormulaTypeEnum.crDateTimeFieldFormatConditionFormulaTypeDateTimeSeparator].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessDateFormat(DateFieldFormat dff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("DateFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("CalendarType", dff.CalendarType.ToStringSafe());
            jsonw.WriteString("DateFirstSeparator", dff.DateFirstSeparator.ToStringSafe());
            jsonw.WriteString("DateOrder", dff.DateOrder.ToStringSafe());
            jsonw.WriteString("DatePrefixSeparator", dff.DatePrefixSeparator.ToStringSafe());
            jsonw.WriteString("DateSecondSeparator", dff.DateSecondSeparator.ToStringSafe());
            jsonw.WriteString("DateSuffixSeparator", dff.DateSuffixSeparator.ToStringSafe());
            jsonw.WriteString("DayFormat", dff.DayFormat.ToStringSafe());
            jsonw.WriteString("DayOfWeekPosition", dff.DayOfWeekPosition.ToStringSafe());
            jsonw.WriteString("DayOfWeekSeparator", dff.DayOfWeekSeparator.ToStringSafe());
            jsonw.WriteString("DayOfWeekType", dff.DayOfWeekType.ToStringSafe());            
            jsonw.WriteString("EraType", dff.EraType.ToStringSafe());
            jsonw.WriteString("MonthFormat", dff.MonthFormat.ToStringSafe());
            jsonw.WriteString("SystemDefaultType", dff.SystemDefaultType.ToStringSafe());
            jsonw.WriteString("YearFormat", dff.YearFormat.ToStringSafe());
            jsonw.WriteString("CalendarTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeCalendarType].Text);
            jsonw.WriteString("DateFirstSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateFirstSeparator].Text);
            jsonw.WriteString("DateOrderFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateOrder].Text);
            jsonw.WriteString("DatePrefixSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDatePrefixSeparator].Text);
            jsonw.WriteString("DateSecondSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateSecondSeparator].Text);
            jsonw.WriteString("DateSuffixSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDateSuffixSeparator].Text);
            jsonw.WriteString("DayFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayFormat].Text);
            jsonw.WriteString("DayOfWeekPositionFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekPosition].Text);
            jsonw.WriteString("DayOfWeekSeparatorFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekSeparator].Text);
            jsonw.WriteString("DayOfWeekTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeDayOfWeekType].Text);
            jsonw.WriteString("EraTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeEraType].Text);
            jsonw.WriteString("MonthFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeMonthFormat].Text);
            jsonw.WriteString("SystemDefaultTypeFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeSystemDefaultType].Text);
            jsonw.WriteString("YearFormatFormula", dff.ConditionFormulas[CrDateFieldFormatConditionFormulaTypeEnum.crDateFieldFormatConditionFormulaTypeYearFormat].Text);          
            jsonw.WriteEndObject();
        }
        
        private static void ProcessNumericFormat(NumericFieldFormat nff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("NumericFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("CurrencyPosition", nff.CurrencyPosition.ToStringSafe());
            jsonw.WriteString("CurrencySymbol", nff.CurrencySymbol);
            jsonw.WriteString("CurrencySymbolFormat", nff.CurrencySymbolFormat.ToStringSafe());
            jsonw.WriteString("DecimalSymbol", nff.DecimalSymbol);
            jsonw.WriteString("DisplayReverseSign", nff.DisplayReverseSign.ToStringSafe());
            jsonw.WriteString("EnableSuppressIfZero", nff.EnableSuppressIfZero.ToStringSafe());
            jsonw.WriteString("EnableUseLeadZero", nff.EnableUseLeadZero.ToStringSafe());
            jsonw.WriteString("NDecimalPlaces", nff.NDecimalPlaces.ToStringSafe());
            jsonw.WriteString("NegativeFormat", nff.NegativeFormat.ToStringSafe());
            jsonw.WriteString("OneCurrencySymbolPerPage", nff.OneCurrencySymbolPerPage.ToStringSafe());
            jsonw.WriteString("RoundingFormat", nff.RoundingFormat.ToStringSafe());
            jsonw.WriteString("ThousandsSeparator", nff.ThousandsSeparator.ToStringSafe());
            jsonw.WriteString("ThousandSymbol", nff.ThousandSymbol);
            jsonw.WriteString("ZeroValueString", nff.ZeroValueString);
            jsonw.WriteString("CurrencyPositionFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencyPosition].Text);
            jsonw.WriteString("CurrencySymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencySymbol].Text);
            jsonw.WriteString("CurrencySymbolFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeCurrencySymbolFormat].Text);
            jsonw.WriteString("DecimalSymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeDecimalSymbol].Text);
            jsonw.WriteString("DisplayReverseSignFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeDisplayReverseSign].Text);
            jsonw.WriteString("EnableSuppressIfZeroFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeEnableSuppressIfZero].Text);
            jsonw.WriteString("EnableUseLeadZeroFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeEnableUseLeadZero].Text);
            jsonw.WriteString("NDecimalPlacesFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeNDecimalPlaces].Text);
            jsonw.WriteString("NegativeFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeNegativeFormat].Text);
            jsonw.WriteString("OneCurrencySymbolPerPageFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeOneCurrencySymbolPerPage].Text);
            jsonw.WriteString("RoundingFormatFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeRoundingFormat].Text);
            jsonw.WriteString("ThousandsSeparatorFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeThousandsSeparator].Text);
            jsonw.WriteString("ThousandSymbolFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeThousandSymbol].Text);
            jsonw.WriteString("ZeroValueStringFormula", nff.ConditionFormulas[CrNumericFieldFormatConditionFormulaTypeEnum.crNumericFieldFormatConditionFormulaTypeZeroValueString].Text);           
            jsonw.WriteEndObject();
        }

        private static void ProcessStringFormat(StringFieldFormat sff, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("StringFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("EnableWordWrap", sff.EnableWordWrap.ToStringSafe());
            jsonw.WriteString("CharacterSpacing", sff.CharacterSpacing.ToStringSafe());
            jsonw.WriteString("MaxNumberOfLines", sff.MaxNumberOfLines.ToStringSafe());
            jsonw.WriteString("ReadingOrder", sff.ReadingOrder.ToStringSafe());
            jsonw.WriteString("TextFormat", sff.TextFormat.ToStringSafe());
            ProcessIndentAndSpacingFormat(sff.IndentAndSpacingFormat,jsonw);
            jsonw.WriteEndObject();
        }

        private static void ProcessIndentAndSpacingFormat(IndentAndSpacingFormat iasf, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("IndentAndSpacingFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("FirstLineIndent", iasf.FirstLineIndent.ToStringSafe());
            jsonw.WriteString("LeftIndent", iasf.LeftIndent.ToStringSafe());
            jsonw.WriteString("LineSpacing", iasf.LineSpacing.ToStringSafe());
            jsonw.WriteString("LineSpacingType", iasf.LineSpacingType.ToStringSafe());
            jsonw.WriteString("RightIndent", iasf.RightIndent.ToStringSafe());
            jsonw.WriteEndObject();
        }

        private static void ProcessTimeFormat(TimeFieldFormat tff, Utf8JsonWriter jsonw)
        {
            
            jsonw.WritePropertyName("TimeFieldFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("AMPMFormat", tff.AMPMFormat.ToStringSafe());
            jsonw.WriteString("AMString", tff.AMString);
            jsonw.WriteString("HourFormat", tff.HourFormat.ToStringSafe());
            jsonw.WriteString("HourMinuteSeparator", tff.HourMinuteSeparator);
            jsonw.WriteString("MinuteFormat", tff.MinuteFormat.ToStringSafe());
            jsonw.WriteString("MinuteSecondSeparator", tff.MinuteSecondSeparator);
            jsonw.WriteString("PMString", tff.PMString);
            jsonw.WriteString("SecondFormat", tff.SecondFormat.ToStringSafe());
            jsonw.WriteString("TimeBase", tff.TimeBase.ToStringSafe());
            jsonw.WriteString("AMPMFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeAMPMFormat].Text);
            jsonw.WriteString("AMStringFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeAMString].Text);
            jsonw.WriteString("HourFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeHourFormat].Text);
            jsonw.WriteString("HourMinuteSeparatorFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeHourMinuteSeparator].Text);
            jsonw.WriteString("MinuteFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeMinuteFormat].Text);
            jsonw.WriteString("MinuteSecondSeparatorFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeMinuteSecondSeparator].Text);
            jsonw.WriteString("PMStringFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypePMString].Text);
            jsonw.WriteString("SecondFormatFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeSecondFormat].Text);
            jsonw.WriteString("TimeBaseFormula", tff.ConditionFormulas[CrTimeFieldFormatConditionFormulaTypeEnum.crTimeFieldFormatConditionFormulaTypeTimeBase].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessBorder(Border border, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("Border");
            jsonw.WriteStartObject();
            jsonw.WriteString("BackgroundColor", border.BackgroundColor.ToStringSafe());
            jsonw.WriteString("BorderColor", border.BorderColor.ToStringSafe());
            jsonw.WriteString("BottomLineStyle", border.BottomLineStyle.ToStringSafe());
            jsonw.WriteString("EnableTightHorizontal", border.EnableTightHorizontal.ToStringSafe());
            jsonw.WriteString("HasDropShadow", border.HasDropShadow.ToStringSafe());
            jsonw.WriteString("LeftLineStyle", border.LeftLineStyle.ToStringSafe());
            jsonw.WriteString("RightLineStyle", border.RightLineStyle.ToStringSafe());
            jsonw.WriteString("TopLineStyle", border.TopLineStyle.ToStringSafe());
            jsonw.WriteString("BackgroundColorFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBackgroundColor].Text);
            jsonw.WriteString("BorderColorFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBorderColor].Text);
            jsonw.WriteString("BottomLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBottomLineStyle].Text);
            jsonw.WriteString("HasDropShadowFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeHasDropShadow].Text);
            jsonw.WriteString("LeftLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeLeftLineStyle].Text);
            jsonw.WriteString("RightLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeRightLineStyle].Text);
            jsonw.WriteString("TightHorizontalFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTightHorizontal].Text);
            jsonw.WriteString("TightVerticalFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTightVertical].Text);
            jsonw.WriteString("TopLineStyleFormula", border.ConditionFormulas[CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeTopLineStyle].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessObjectFormat(ObjectFormat of, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("ObjectFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("CssClass", of.CssClass);
            jsonw.WriteString("EnableCanGrow", of.EnableCanGrow.ToStringSafe());
            jsonw.WriteString("EnableCloseAtPageBreak", of.EnableCloseAtPageBreak.ToStringSafe());
            jsonw.WriteString("EnableKeepTogether", of.EnableKeepTogether.ToStringSafe());
            jsonw.WriteString("EnableSuppress", of.EnableSuppress.ToStringSafe());
            jsonw.WriteString("HorizontalAlignment", of.HorizontalAlignment.ToStringSafe());
            jsonw.WriteString("HyperlinkText", of.HyperlinkText);
            jsonw.WriteString("HyperlinkType", of.HyperlinkType.ToStringSafe());
            jsonw.WriteString("TextRotationAngle", of.TextRotationAngle.ToStringSafe());
            jsonw.WriteString("ToolTipText", of.ToolTipText);
            jsonw.WriteString("VerticalAlignment", of.VerticalAlignment.ToStringSafe());
            jsonw.WriteString("CssClassFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeCssClass].Text);
            jsonw.WriteString("DeltaWidthFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaWidth].Text);
            jsonw.WriteString("DeltaXFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDeltaX].Text);
            jsonw.WriteString("DisplayStringFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDisplayString].Text);
            jsonw.WriteString("EnableCanGrowFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCanGrow].Text);
            jsonw.WriteString("EnableCloseAtPageBreakFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableCloseAtPageBreak].Text);
            jsonw.WriteString("EnableKeepTogetherFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableKeepTogether].Text);
            jsonw.WriteString("EnableSuppressFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableSuppress].Text);
            jsonw.WriteString("HorizontalAlignmentFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHorizontalAlignment].Text);
            jsonw.WriteString("HyperlinkFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeHyperlink].Text);
            jsonw.WriteString("TextRotationAngleFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeRotation].Text);
            jsonw.WriteString("ToolTipTextFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeToolTipText].Text);
            jsonw.WriteString("VerticalAlignmentFormula", of.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeVerticalAlignment].Text);
            jsonw.WriteEndObject();
        }

        private static void ProcessSectionFormat(SectionFormat sf, Utf8JsonWriter jsonw)
        {
            jsonw.WritePropertyName("SectionFormat");
            jsonw.WriteStartObject();
            jsonw.WriteString("BackgroundColor", sf.BackgroundColor.ToStringSafe());
            jsonw.WriteString("CssClass", sf.CssClass);
            jsonw.WriteString("EnableKeepTogether", sf.EnableKeepTogether.ToStringSafe());
            jsonw.WriteString("EnableNewPageAfter", sf.EnableNewPageAfter.ToStringSafe());
            jsonw.WriteString("EnableNewPageBefore", sf.EnableNewPageBefore.ToStringSafe());
            jsonw.WriteString("EnablePrintAtBottomOfPage", sf.EnablePrintAtBottomOfPage.ToStringSafe());
            jsonw.WriteString("EnableResetPageNumberAfter", sf.EnableResetPageNumberAfter.ToStringSafe());
            jsonw.WriteString("EnableSuppress", sf.EnableSuppress.ToStringSafe());
            jsonw.WriteString("EnableSuppressIfBlank", sf.EnableSuppressIfBlank.ToStringSafe());
            jsonw.WriteString("EnableUnderlaySection", sf.EnableUnderlaySection.ToStringSafe());
            jsonw.WriteString("PageOrientation", sf.PageOrientation.ToStringSafe());
            jsonw.WriteString("BackgroundColorFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text);
            jsonw.WriteString("EnableClampPageFooterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text);
            jsonw.WriteString("EnableHideForDrillDownFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text);
            jsonw.WriteString("EnableKeepTogetherFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text);
            jsonw.WriteString("EnableNewPageAfterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text);
            jsonw.WriteString("EnableNewPageBeforeFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text);
            jsonw.WriteString("EnablePrintAtBottomOfPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text);
            jsonw.WriteString("EnableResetPageNumberAfterFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text);
            jsonw.WriteString("EnableSuppressFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text);
            jsonw.WriteString("EnableSuppressIfBlankFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text);
            jsonw.WriteString("EnableUnderlaySectionFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text);
            jsonw.WriteString("GroupNumberPerPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text);
            jsonw.WriteString("RecordNumberPerPageFormula", sf.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text);
            jsonw.WriteEndObject();
        }
        
    }
}
