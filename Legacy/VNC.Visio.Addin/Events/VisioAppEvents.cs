using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//using moiVisio = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Visio;
//using VisioHelper = VNC.Visio.AddinHelper.Visio;
using VNC;
//using SupportTools_Visio.Core;

namespace VNC.Visio.Addin.Events
{
    public class VisioAppEvents
    {
        private Application _VisioApplication;
        public Application VisioApplication
        {
            get
            {
                return _VisioApplication;
            }
            set
            {
                if (_VisioApplication != null)
                {
                    // Should remove all the event handlers;
                }

                _VisioApplication = value;

                if (_VisioApplication != null)
                {
                    _VisioApplication.AfterModal += new EApplication_AfterModalEventHandler(_VisioApplication_AfterModal);
                    _VisioApplication.AfterRemoveHiddenInformation += new EApplication_AfterRemoveHiddenInformationEventHandler(_VisioApplication_AfterRemoveHiddenInformation);
                    _VisioApplication.AfterResume += new EApplication_AfterResumeEventHandler(_VisioApplication_AfterResume);
                    _VisioApplication.AfterResumeEvents += new EApplication_AfterResumeEventsEventHandler(_VisioApplication_AfterResumeEvents);
                    _VisioApplication.AppActivated += new EApplication_AppActivatedEventHandler(_VisioApplication_AppActivated);
                    _VisioApplication.AppDeactivated += new EApplication_AppDeactivatedEventHandler(_VisioApplication_AppDeactivated);
                    _VisioApplication.AppObjActivated += new EApplication_AppObjActivatedEventHandler(_VisioApplication_AppObjActivated);
                    _VisioApplication.AppObjDeactivated += new EApplication_AppObjDeactivatedEventHandler(_VisioApplication_AppObjDeactivated);
                    _VisioApplication.BeforeDataRecordsetDelete += new EApplication_BeforeDataRecordsetDeleteEventHandler(_VisioApplication_BeforeDataRecordsetDelete);
                    _VisioApplication.BeforeDocumentClose += new EApplication_BeforeDocumentCloseEventHandler(_VisioApplication_BeforeDocumentClose);
                    _VisioApplication.BeforeDocumentSave += new EApplication_BeforeDocumentSaveEventHandler(_VisioApplication_BeforeDocumentSave);
                    _VisioApplication.BeforeDocumentSaveAs += new EApplication_BeforeDocumentSaveAsEventHandler(_VisioApplication_BeforeDocumentSaveAs);
                    _VisioApplication.BeforeMasterDelete += new EApplication_BeforeMasterDeleteEventHandler(_VisioApplication_BeforeMasterDelete);
                    _VisioApplication.BeforeModal += new EApplication_BeforeModalEventHandler(_VisioApplication_BeforeModal);
                    _VisioApplication.BeforePageDelete += new EApplication_BeforePageDeleteEventHandler(_VisioApplication_BeforePageDelete);
                    _VisioApplication.BeforeQuit += new EApplication_BeforeQuitEventHandler(_VisioApplication_BeforeQuit);
                    _VisioApplication.BeforeSelectionDelete += new EApplication_BeforeSelectionDeleteEventHandler(_VisioApplication_BeforeSelectionDelete);
                    _VisioApplication.BeforeShapeDelete += new EApplication_BeforeShapeDeleteEventHandler(_VisioApplication_BeforeShapeDelete);
                    _VisioApplication.BeforeShapeTextEdit += new EApplication_BeforeShapeTextEditEventHandler(_VisioApplication_BeforeShapeTextEdit);
                    _VisioApplication.BeforeStyleDelete += new EApplication_BeforeStyleDeleteEventHandler(_VisioApplication_BeforeStyleDelete);
                    _VisioApplication.BeforeSuspend += new EApplication_BeforeSuspendEventHandler(_VisioApplication_BeforeSuspend);
                    _VisioApplication.BeforeSuspendEvents += new EApplication_BeforeSuspendEventsEventHandler(_VisioApplication_BeforeSuspendEvents);
                    _VisioApplication.BeforeWindowClosed += new EApplication_BeforeWindowClosedEventHandler(_VisioApplication_BeforeWindowClosed);
                    _VisioApplication.BeforeWindowPageTurn += new EApplication_BeforeWindowPageTurnEventHandler(_VisioApplication_BeforeWindowPageTurn);
                    _VisioApplication.BeforeWindowSelDelete += new EApplication_BeforeWindowSelDeleteEventHandler(_VisioApplication_BeforeWindowSelDelete);
                    _VisioApplication.CalloutRelationshipAdded += new EApplication_CalloutRelationshipAddedEventHandler(_VisioApplication_CalloutRelationshipAdded);
                    _VisioApplication.CalloutRelationshipDeleted += new EApplication_CalloutRelationshipDeletedEventHandler(_VisioApplication_CalloutRelationshipDeleted);
                    _VisioApplication.CellChanged += new EApplication_CellChangedEventHandler(_VisioApplication_CellChanged);
                    _VisioApplication.ConnectionsAdded += new EApplication_ConnectionsAddedEventHandler(_VisioApplication_ConnectionsAdded);
                    _VisioApplication.ConnectionsDeleted += new EApplication_ConnectionsDeletedEventHandler(_VisioApplication_ConnectionsDeleted);
                    _VisioApplication.ContainerRelationshipAdded += new EApplication_ContainerRelationshipAddedEventHandler(_VisioApplication_ContainerRelationshipAdded);
                    _VisioApplication.ContainerRelationshipDeleted += new EApplication_ContainerRelationshipDeletedEventHandler(_VisioApplication_ContainerRelationshipDeleted);
                    _VisioApplication.ConvertToGroupCanceled += new EApplication_ConvertToGroupCanceledEventHandler(_VisioApplication_ConvertToGroupCanceled);
                    _VisioApplication.DataRecordsetAdded += new EApplication_DataRecordsetAddedEventHandler(_VisioApplication_DataRecordsetAdded);
                    _VisioApplication.DataRecordsetChanged += new EApplication_DataRecordsetChangedEventHandler(_VisioApplication_DataRecordsetChanged);
                    _VisioApplication.DesignModeEntered += new EApplication_DesignModeEnteredEventHandler(_VisioApplication_DesignModeEntered);
                    _VisioApplication.DocumentChanged += new EApplication_DocumentChangedEventHandler(_VisioApplication_DocumentChanged);
                    _VisioApplication.DocumentCloseCanceled += new EApplication_DocumentCloseCanceledEventHandler(_VisioApplication_DocumentCloseCanceled);
                    _VisioApplication.DocumentCreated += new EApplication_DocumentCreatedEventHandler(_VisioApplication_DocumentCreated);
                    _VisioApplication.DocumentOpened += new EApplication_DocumentOpenedEventHandler(_VisioApplication_DocumentOpened);
                    _VisioApplication.DocumentSaved += new EApplication_DocumentSavedEventHandler(_VisioApplication_DocumentSaved);
                    _VisioApplication.DocumentSavedAs += new EApplication_DocumentSavedAsEventHandler(_VisioApplication_DocumentSavedAs);
                    _VisioApplication.EnterScope += new EApplication_EnterScopeEventHandler(_VisioApplication_EnterScope);
                    _VisioApplication.ExitScope += new EApplication_ExitScopeEventHandler(_VisioApplication_ExitScope);
                    _VisioApplication.FormulaChanged += new EApplication_FormulaChangedEventHandler(_VisioApplication_FormulaChanged);
                    _VisioApplication.GroupCanceled += new EApplication_GroupCanceledEventHandler(_VisioApplication_GroupCanceled);
                    _VisioApplication.KeyDown += new EApplication_KeyDownEventHandler(_VisioApplication_KeyDown);
                    _VisioApplication.KeyPress += new EApplication_KeyPressEventHandler(_VisioApplication_KeyPress);
                    _VisioApplication.KeyUp += new EApplication_KeyUpEventHandler(_VisioApplication_KeyUp);
                    _VisioApplication.MarkerEvent += new EApplication_MarkerEventEventHandler(_VisioApplication_MarkerEvent);
                    _VisioApplication.MasterAdded += new EApplication_MasterAddedEventHandler(_VisioApplication_MasterAdded);
                    _VisioApplication.MasterChanged += new EApplication_MasterChangedEventHandler(_VisioApplication_MasterChanged);
                    _VisioApplication.MasterDeleteCanceled += new EApplication_MasterDeleteCanceledEventHandler(_VisioApplication_MasterDeleteCanceled);
                    _VisioApplication.MouseDown += new EApplication_MouseDownEventHandler(_VisioApplication_MouseDown);
                    _VisioApplication.MouseMove += new EApplication_MouseMoveEventHandler(_VisioApplication_MouseMove);
                    _VisioApplication.MouseUp += new EApplication_MouseUpEventHandler(_VisioApplication_MouseUp);
                    _VisioApplication.MustFlushScopeBeginning += new EApplication_MustFlushScopeBeginningEventHandler(_VisioApplication_MustFlushScopeBeginning);
                    _VisioApplication.MustFlushScopeEnded += new EApplication_MustFlushScopeEndedEventHandler(_VisioApplication_MustFlushScopeEnded);
                    _VisioApplication.NoEventsPending += new EApplication_NoEventsPendingEventHandler(_VisioApplication_NoEventsPending);
                    _VisioApplication.OnKeystrokeMessageForAddon += new EApplication_OnKeystrokeMessageForAddonEventHandler(_VisioApplication_OnKeystrokeMessageForAddon);
                    _VisioApplication.PageAdded += new EApplication_PageAddedEventHandler(_VisioApplication_PageAdded);
                    _VisioApplication.PageChanged += new EApplication_PageChangedEventHandler(_VisioApplication_PageChanged);
                    _VisioApplication.PageDeleteCanceled += new EApplication_PageDeleteCanceledEventHandler(_VisioApplication_PageDeleteCanceled);
                    _VisioApplication.QueryCancelConvertToGroup += new EApplication_QueryCancelConvertToGroupEventHandler(_VisioApplication_QueryCancelConvertToGroup);
                    _VisioApplication.QueryCancelDocumentClose += new EApplication_QueryCancelDocumentCloseEventHandler(_VisioApplication_QueryCancelDocumentClose);
                    _VisioApplication.QueryCancelGroup += new EApplication_QueryCancelGroupEventHandler(_VisioApplication_QueryCancelGroup);
                    _VisioApplication.QueryCancelMasterDelete += new EApplication_QueryCancelMasterDeleteEventHandler(_VisioApplication_QueryCancelMasterDelete);
                    _VisioApplication.QueryCancelPageDelete += new EApplication_QueryCancelPageDeleteEventHandler(_VisioApplication_QueryCancelPageDelete);
                    _VisioApplication.QueryCancelQuit += new EApplication_QueryCancelQuitEventHandler(_VisioApplication_QueryCancelQuit);
                    _VisioApplication.QueryCancelSelectionDelete += new EApplication_QueryCancelSelectionDeleteEventHandler(_VisioApplication_QueryCancelSelectionDelete);
                    _VisioApplication.QueryCancelStyleDelete += new EApplication_QueryCancelStyleDeleteEventHandler(_VisioApplication_QueryCancelStyleDelete);
                    _VisioApplication.QueryCancelSuspend += new EApplication_QueryCancelSuspendEventHandler(_VisioApplication_QueryCancelSuspend);
                    _VisioApplication.QueryCancelSuspendEvents += new EApplication_QueryCancelSuspendEventsEventHandler(_VisioApplication_QueryCancelSuspendEvents);
                    _VisioApplication.QueryCancelUngroup += new EApplication_QueryCancelUngroupEventHandler(_VisioApplication_QueryCancelUngroup);
                    _VisioApplication.QueryCancelWindowClose += new EApplication_QueryCancelWindowCloseEventHandler(_VisioApplication_QueryCancelWindowClose);
                    _VisioApplication.QuitCanceled += new EApplication_QuitCanceledEventHandler(_VisioApplication_QuitCanceled);
                    _VisioApplication.RuleSetValidated += new EApplication_RuleSetValidatedEventHandler(_VisioApplication_RuleSetValidated);
                    _VisioApplication.RunModeEntered += new EApplication_RunModeEnteredEventHandler(_VisioApplication_RunModeEntered);
                    _VisioApplication.SelectionAdded += new EApplication_SelectionAddedEventHandler(_VisioApplication_SelectionAdded);
                    _VisioApplication.SelectionChanged += new EApplication_SelectionChangedEventHandler(_VisioApplication_SelectionChanged);
                    _VisioApplication.SelectionDeleteCanceled += new EApplication_SelectionDeleteCanceledEventHandler(_VisioApplication_SelectionDeleteCanceled);
                    _VisioApplication.ShapeAdded += new EApplication_ShapeAddedEventHandler(_VisioApplication_ShapeAdded);
                    _VisioApplication.ShapeChanged += new EApplication_ShapeChangedEventHandler(_VisioApplication_ShapeChanged);
                    _VisioApplication.ShapeDataGraphicChanged += new EApplication_ShapeDataGraphicChangedEventHandler(_VisioApplication_ShapeDataGraphicChanged);
                    _VisioApplication.ShapeExitedTextEdit += new EApplication_ShapeExitedTextEditEventHandler(_VisioApplication_ShapeExitedTextEdit);
                    _VisioApplication.ShapeLinkAdded += new EApplication_ShapeLinkAddedEventHandler(_VisioApplication_ShapeLinkAdded);
                    _VisioApplication.ShapeLinkDeleted += new EApplication_ShapeLinkDeletedEventHandler(_VisioApplication_ShapeLinkDeleted);
                    _VisioApplication.ShapeParentChanged += new EApplication_ShapeParentChangedEventHandler(_VisioApplication_ShapeParentChanged);
                    _VisioApplication.StyleAdded += new EApplication_StyleAddedEventHandler(_VisioApplication_StyleAdded);
                    _VisioApplication.StyleChanged += new EApplication_StyleChangedEventHandler(_VisioApplication_StyleChanged);
                    _VisioApplication.StyleDeleteCanceled += new EApplication_StyleDeleteCanceledEventHandler(_VisioApplication_StyleDeleteCanceled);
                    _VisioApplication.SuspendCanceled += new EApplication_SuspendCanceledEventHandler(_VisioApplication_SuspendCanceled);
                    _VisioApplication.SuspendEventsCanceled += new EApplication_SuspendEventsCanceledEventHandler(_VisioApplication_SuspendEventsCanceled);
                    _VisioApplication.TextChanged += new EApplication_TextChangedEventHandler(_VisioApplication_TextChanged);
                    _VisioApplication.UngroupCanceled += new EApplication_UngroupCanceledEventHandler(_VisioApplication_UngroupCanceled);
                    _VisioApplication.ViewChanged += new EApplication_ViewChangedEventHandler(_VisioApplication_ViewChanged);
                    _VisioApplication.WindowActivated += new EApplication_WindowActivatedEventHandler(_VisioApplication_WindowActivated);
                    _VisioApplication.WindowChanged += new EApplication_WindowChangedEventHandler(_VisioApplication_WindowChanged);
                    _VisioApplication.WindowCloseCanceled += new EApplication_WindowCloseCanceledEventHandler(_VisioApplication_WindowCloseCanceled);
                    _VisioApplication.WindowOpened += new EApplication_WindowOpenedEventHandler(_VisioApplication_WindowOpened);
                    _VisioApplication.WindowTurnedToPage += new EApplication_WindowTurnedToPageEventHandler(_VisioApplication_WindowTurnedToPage);
                }
            }
        }

        #region Events that Do Something and Log

        short countWindowTurnedToPage;
        void _VisioApplication_WindowTurnedToPage(Window Window)
        {
            DisplayInWatchWindow(countWindowTurnedToPage++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            Window.ViewFit = (int)VisWindowFit.visFitPage;
        }

        short countShapeAdded;
        void _VisioApplication_ShapeAdded(Shape Shape)
        {
            DisplayInWatchWindow(countShapeAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            Actions.Visio_Shape.HandleShapeAdded(Shape);
        }

        short countPageChanged;
        void _VisioApplication_PageChanged(Page Page)
        {
            DisplayInWatchWindow(countPageChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);

            Actions.Visio_Page.PageChanged(Page);
        }

        short countMarkerEvent;
        void _VisioApplication_MarkerEvent(Application app, int SequenceNum, string ContextString)
        {
            string message = string.Format("{0}  SequenceNum={1}  ContextString=>{2}<",
                System.Reflection.MethodInfo.GetCurrentMethod().Name,
                SequenceNum,
                ContextString);
            DisplayInWatchWindow(countMarkerEvent++, message);

            // If we got here from a RUNADDONWARGS("QueueMarkerEvent", "<Action>")
            // the ContextString should have multiple pieces showing the context of what was selected.
            try
            {
                if (null != ContextString)
                {
                    var context = ContextString.Split(' ');

                    if (context.Count() > 1)
                    {
                        RouteShapeSheet_QueueMarkerEvent(app, SequenceNum, context); ;
                    }
                    else
                    {
                        // Quietly ignore
                    }
                }
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        #endregion

        #region Events and Chatty Events that just Log

        #region Regular Events - Just Log

        short countBeforeMasterDelete;
        void _VisioApplication_BeforeMasterDelete(Master Master)
        {
            DisplayInWatchWindow(countBeforeMasterDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterDeleteCanceled;
        void _VisioApplication_MasterDeleteCanceled(Master Master)
        {
            DisplayInWatchWindow(countMasterDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterChanged;
        void _VisioApplication_MasterChanged(Master Master)
        {
            DisplayInWatchWindow(countMasterChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countMasterAdded;
        void _VisioApplication_MasterAdded(Master Master)
        {
            DisplayInWatchWindow(countMasterAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countTextChange;
        void _VisioApplication_TextChanged(Shape Shape)
        {
            DisplayInWatchWindow(countTextChange++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowOpened;
        void _VisioApplication_WindowOpened(Window Window)
        {
            DisplayInWatchWindow(countWindowOpened++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowCloseCanceled;
        void _VisioApplication_WindowCloseCanceled(Window Window)
        {
            DisplayInWatchWindow(countWindowCloseCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowChanged;
        void _VisioApplication_WindowChanged(Window Window)
        {
            DisplayInWatchWindow(countWindowChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countWindowActivated;
        void _VisioApplication_WindowActivated(Window Window)
        {
            DisplayInWatchWindow(countWindowActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countViewChanged;
        void _VisioApplication_ViewChanged(Window Window)
        {
            DisplayInWatchWindow(countViewChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countUngroupCanceled;
        void _VisioApplication_UngroupCanceled(Selection Selection)
        {
            DisplayInWatchWindow(countUngroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countSuspendEventsCanceled;
        void _VisioApplication_SuspendEventsCanceled(Application app)
        {
            DisplayInWatchWindow(countSuspendEventsCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countSuspendCanceled;
        void _VisioApplication_SuspendCanceled(Application app)
        {
            DisplayInWatchWindow(countSuspendCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleDeleteCanceled;
        void _VisioApplication_StyleDeleteCanceled(Style Style)
        {
            DisplayInWatchWindow(countStyleDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleChanged;
        void _VisioApplication_StyleChanged(Style Style)
        {
            DisplayInWatchWindow(countStyleChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countStyleAdded;
        void _VisioApplication_StyleAdded(Style Style)
        {
            DisplayInWatchWindow(countStyleAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeParentChanged;
        void _VisioApplication_ShapeParentChanged(Shape Shape)
        {
            DisplayInWatchWindow(countShapeParentChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeLinkDeleted;
        void _VisioApplication_ShapeLinkDeleted(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayInWatchWindow(countShapeLinkDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeLinkAdded;
        void _VisioApplication_ShapeLinkAdded(Shape Shape, int DataRecordsetID, int DataRowID)
        {
            DisplayInWatchWindow(countShapeLinkAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeExitedTextEdit;
        void _VisioApplication_ShapeExitedTextEdit(Shape Shape)
        {
            DisplayInWatchWindow(countShapeExitedTextEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeDataGraphicChanged;
        void _VisioApplication_ShapeDataGraphicChanged(Shape Shape)
        {
            DisplayInWatchWindow(countShapeDataGraphicChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countShapeChanged;
        void _VisioApplication_ShapeChanged(Shape Shape)
        {
            DisplayInWatchWindow(countShapeChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionDeleteCanceled;
        void _VisioApplication_SelectionDeleteCanceled(Selection Selection)
        {
            DisplayInWatchWindow(countSelectionDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionChanged;
        void _VisioApplication_SelectionChanged(Window Window)
        {
            Common.EventAggregator.GetEvent<SelectionChangedEvent>().Publish();
            DisplayInWatchWindow(countSelectionChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countSelectionAdded;
        void _VisioApplication_SelectionAdded(Selection Selection)
        {
            DisplayInWatchWindow(countSelectionAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countRunModeEntered;
        void _VisioApplication_RunModeEntered(Document Doc)
        {
            DisplayInWatchWindow(countRunModeEntered++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countRuleSetValidated;
        void _VisioApplication_RuleSetValidated(ValidationRuleSet RuleSet)
        {
            DisplayInWatchWindow(countRuleSetValidated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countQuitCanceled;
        void _VisioApplication_QuitCanceled(Application app)
        {
            DisplayInWatchWindow(countQuitCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countQueryCancelWindowClose;
        bool _VisioApplication_QueryCancelWindowClose(Window Window)
        {
            DisplayInWatchWindow(countQueryCancelWindowClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelUngroup;
        bool _VisioApplication_QueryCancelUngroup(Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelUngroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSuspendEvents;
        bool _VisioApplication_QueryCancelSuspendEvents(Application app)
        {
            DisplayInWatchWindow(countQueryCancelSuspendEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSuspend;
        bool _VisioApplication_QueryCancelSuspend(Application app)
        {
            DisplayInWatchWindow(countQueryCancelSuspend++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelStyleDelete;
        bool _VisioApplication_QueryCancelStyleDelete(Style Style)
        {
            DisplayInWatchWindow(countQueryCancelStyleDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelSelectionDelete;
        bool _VisioApplication_QueryCancelSelectionDelete(Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelSelectionDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelQuit;
        bool _VisioApplication_QueryCancelQuit(Application app)
        {
            DisplayInWatchWindow(countQueryCancelQuit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelPageDelete;
        bool _VisioApplication_QueryCancelPageDelete(Page Page)
        {
            DisplayInWatchWindow(countQueryCancelPageDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelMasterDelete;
        bool _VisioApplication_QueryCancelMasterDelete(Master Master)
        {
            DisplayInWatchWindow(countQueryCancelMasterDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelGroup;
        bool _VisioApplication_QueryCancelGroup(Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelGroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelDocumentClose;
        bool _VisioApplication_QueryCancelDocumentClose(Document Doc)
        {
            DisplayInWatchWindow(countQueryCancelDocumentClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countQueryCancelConvertToGroup;
        bool _VisioApplication_QueryCancelConvertToGroup(Selection Selection)
        {
            DisplayInWatchWindow(countQueryCancelConvertToGroup++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countPageDeleteCanceled;
        void _VisioApplication_PageDeleteCanceled(Page Page)
        {
            DisplayInWatchWindow(countPageDeleteCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countPageAdded;
        void _VisioApplication_PageAdded(Page Page)
        {
            DisplayInWatchWindow(countPageAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countOnKeystrokeMessageForAddon;
        bool _VisioApplication_OnKeystrokeMessageForAddon(MSGWrap MSG)
        {
            DisplayInWatchWindow(countOnKeystrokeMessageForAddon++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            return false;
        }

        short countGroupCanceled;
        void _VisioApplication_GroupCanceled(Selection Selection)
        {
            DisplayInWatchWindow(countGroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countFormulaChanged;
        void _VisioApplication_FormulaChanged(Cell Cell)
        {
            DisplayInWatchWindow(countFormulaChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countExitScope;
        void _VisioApplication_ExitScope(Application app, int nScopeID, string bstrDescription, bool bErrOrCancelled)
        {
            DisplayInWatchWindow(countExitScope++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countEnterScope;
        void _VisioApplication_EnterScope(Application app, int nScopeID, string bstrDescription)
        {
            DisplayInWatchWindow(countEnterScope++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentSavedAs;
        void _VisioApplication_DocumentSavedAs(Document Doc)
        {
            DisplayInWatchWindow(countDocumentSavedAs++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentSaved;
        void _VisioApplication_DocumentSaved(Document Doc)
        {
            DisplayInWatchWindow(countDocumentSaved++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentOpened;
        void _VisioApplication_DocumentOpened(Document Doc)
        {
            DisplayInWatchWindow(countDocumentOpened++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentCreated;
        void _VisioApplication_DocumentCreated(Document Doc)
        {
            DisplayInWatchWindow(countDocumentCreated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentCloseCanceled;
        void _VisioApplication_DocumentCloseCanceled(Document Doc)
        {
            DisplayInWatchWindow(countDocumentCloseCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDocumentChanged;
        void _VisioApplication_DocumentChanged(Document Doc)
        {
            DisplayInWatchWindow(countDocumentChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDesignModeEntered;
        void _VisioApplication_DesignModeEntered(Document Doc)
        {
            DisplayInWatchWindow(countDesignModeEntered++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDataRecordsetChanged;
        void _VisioApplication_DataRecordsetChanged(DataRecordsetChangedEvent DataRecordsetChanged)
        {
            DisplayInWatchWindow(countDataRecordsetChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countDataRecordsetAdded;
        void _VisioApplication_DataRecordsetAdded(DataRecordset DataRecordset)
        {
            DisplayInWatchWindow(countDataRecordsetAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countConvertToGroupCanceled;
        void _VisioApplication_ConvertToGroupCanceled(Selection Selection)
        {
            DisplayInWatchWindow(countConvertToGroupCanceled++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countContainerRelationshipDeleted;
        void _VisioApplication_ContainerRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countContainerRelationshipDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countContainerRelationshipAdded;
        void _VisioApplication_ContainerRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countContainerRelationshipAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countConnectionsDeleted;
        void _VisioApplication_ConnectionsDeleted(Connects Connects)
        {
            DisplayInWatchWindow(countConnectionsDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countConnectionsAdded;
        void _VisioApplication_ConnectionsAdded(Connects Connects)
        {
            DisplayInWatchWindow(countConnectionsAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countCellChanged;
        void _VisioApplication_CellChanged(Cell Cell)
        {
            DisplayInWatchWindow(countCellChanged++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countCalloutRelationshipDeleted;
        void _VisioApplication_CalloutRelationshipDeleted(RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countCalloutRelationshipDeleted++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countCalloutRelationshipAdded;
        void _VisioApplication_CalloutRelationshipAdded(RelatedShapePairEvent ShapePair)
        {
            DisplayInWatchWindow(countCalloutRelationshipAdded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowSelDelete;
        void _VisioApplication_BeforeWindowSelDelete(Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowSelDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowPageTurn;
        void _VisioApplication_BeforeWindowPageTurn(Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowPageTurn++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeWindowClosed;
        void _VisioApplication_BeforeWindowClosed(Window Window)
        {
            DisplayInWatchWindow(countBeforeWindowClosed++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSuspendEvents;
        void _VisioApplication_BeforeSuspendEvents(Application app)
        {
            DisplayInWatchWindow(countBeforeSuspendEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSuspend;
        void _VisioApplication_BeforeSuspend(Application app)
        {
            DisplayInWatchWindow(countBeforeSuspend++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeStyleDelete;
        void _VisioApplication_BeforeStyleDelete(Style Style)
        {
            DisplayInWatchWindow(countBeforeStyleDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeShapeTextEdit;
        void _VisioApplication_BeforeShapeTextEdit(Shape Shape)
        {
            DisplayInWatchWindow(countBeforeShapeTextEdit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeShapeDelete;
        void _VisioApplication_BeforeShapeDelete(Shape Shape)
        {
            DisplayInWatchWindow(countBeforeShapeDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeSelectionDelete;
        void _VisioApplication_BeforeSelectionDelete(Selection Selection)
        {
            DisplayInWatchWindow(countBeforeSelectionDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeQuit;
        void _VisioApplication_BeforeQuit(Application app)
        {
            DisplayInWatchWindow(countBeforeQuit++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforePageDelete;
        void _VisioApplication_BeforePageDelete(Page Page)
        {
            DisplayInWatchWindow(countBeforePageDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeModal;
        void _VisioApplication_BeforeModal(Application app)
        {
            DisplayInWatchWindow(countBeforeModal++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentSaveAs;
        void _VisioApplication_BeforeDocumentSaveAs(Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentSaveAs++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentSave;
        void _VisioApplication_BeforeDocumentSave(Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentSave++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDocumentClose;
        void _VisioApplication_BeforeDocumentClose(Document Doc)
        {
            DisplayInWatchWindow(countBeforeDocumentClose++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countBeforeDataRecordsetDelete;
        void _VisioApplication_BeforeDataRecordsetDelete(DataRecordset DataRecordset)
        {
            DisplayInWatchWindow(countBeforeDataRecordsetDelete++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAppObjDeactivated;
        void _VisioApplication_AppObjDeactivated(Application app)
        {
            DisplayInWatchWindow(countAppObjDeactivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAppObjActivated;
        void _VisioApplication_AppObjActivated(Application app)
        {
            DisplayInWatchWindow(countAppObjActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAppDeactivated;
        void _VisioApplication_AppDeactivated(Application app)
        {
            DisplayInWatchWindow(countAppDeactivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAppActivated;
        void _VisioApplication_AppActivated(Application app)
        {
            DisplayInWatchWindow(countAppActivated++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterResumeEvents;
        void _VisioApplication_AfterResumeEvents(Application app)
        {
            DisplayInWatchWindow(countAfterResumeEvents++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterResume;
        void _VisioApplication_AfterResume(Application app)
        {
            DisplayInWatchWindow(countAfterResume++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterRemoveHiddenInformation;
        void _VisioApplication_AfterRemoveHiddenInformation(Document Doc)
        {
            DisplayInWatchWindow(countAfterRemoveHiddenInformation++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        short countAfterModal;
        void _VisioApplication_AfterModal(Application app)
        {
            DisplayInWatchWindow(countAfterModal++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        #endregion

        #region Chatty Events

        short countKeyUp;
        void _VisioApplication_KeyUp(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countKeyUp++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyUp++;
            }
        }

        short countKeyPress;
        void _VisioApplication_KeyPress(int KeyAscii, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countKeyPress++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyPress++;
            }
        }

        short countKeyDown;
        void _VisioApplication_KeyDown(int KeyCode, int KeyButtonState, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countKeyDown++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countKeyDown++;
            }
        }
        short countNoEventsPending;
        void _VisioApplication_NoEventsPending(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countNoEventsPending++, System.Reflection.MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countNoEventsPending++;
            }
        }

        short countMustFlushScopeEnded;
        void _VisioApplication_MustFlushScopeEnded(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMustFlushScopeEnded++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countMustFlushScopeEnded++;
            }
        }

        short countMustFlushScopeBeginning;
        void _VisioApplication_MustFlushScopeBeginning(Application app)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMustFlushScopeBeginning++, System.Reflection.MethodInfo.GetCurrentMethod().Name);
            }
            else
            {
                countMustFlushScopeBeginning++;
            }
        }

        short countMouseDown;
        void _VisioApplication_MouseDown(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseDown++, System.Reflection.MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseDown++;
            }
        }

        short countMouseUp;
        void _VisioApplication_MouseUp(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseUp++, System.Reflection.MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseUp++;
            }
        }

        short countMouseMove;
        void _VisioApplication_MouseMove(int Button, int KeyButtonState, double x, double y, ref bool CancelDefault)
        {
            if (Common.DisplayChattyEvents)
            {
                DisplayInWatchWindow(countMouseMove++, System.Reflection.MethodInfo.GetCurrentMethod().Name); ;
            }
            else
            {
                countMouseMove++;
            }
        }

        #endregion Chatty Events

        #endregion Events and Chatty Events that just Log

        private void DisplayInWatchWindow(short i, string outputLine)
        {
            if (Common.DisplayEvents)
            {
                VNC.Visio.AddinHelper.Common.WriteToWatchWindow(string.Format("{0}:{1}", outputLine, i));
            }
        }

        private void RouteShapeSheet_QueueMarkerEvent(Application app, int sequenceNum, string[] context)
        {
            VNC.Log.Debug("", Common.LOG_CATEGORY, 0);

            VisioHelper.DisplayInWatchWindow(string.Format("{0}()",
                System.Reflection.MethodInfo.GetCurrentMethod().Name));

            try
            {
                for (int i = 0; i < context.Count(); i++)
                {
                    VisioHelper.DisplayInWatchWindow(string.Format("  ci[{0}]:>{1}", i, context[i]));
                }

                // The QueueMarkerEvent provides context information for each event along with user information (action).
                // Each part of the context is preceeded by an identifier of the form /<identifier>=
                // Grab the part of the entry that past the = sign.

                string doc = context[0].Substring(5);       // "/doc="
                string page = context[1].Substring(6);      // "/page="
                string shape = context[2].Substring(7);     // "/shape="

                VisioHelper.DisplayInWatchWindow(string.Format("   doc:   >{0}<", doc));
                VisioHelper.DisplayInWatchWindow(string.Format("   page:  >{0}<", page));
                VisioHelper.DisplayInWatchWindow(string.Format("   shape: >{0}<", shape));

                // QueueMarkerEvent from Pages does not have a shapeu

                string shapeu = "<none>";

                if (context.Count() > 3)
                {
                    shapeu = context[3].Substring(8);    // "/shapeu="
                    DisplayInWatchWindow(0, string.Format("   shapeu:>{0}<", shapeu));
                }

                string args = context[4].Replace("%20", " ");   // Embedded spaces
                var actionArgs = args.Split(',');

                DisplayInWatchWindow(0, string.Format("    actionArgs:>{0}<", actionArgs[0]));

                // TODO:
                // Add new case statement for each unique "<Action>"
                // RUNADDONWARGS("QueueMarkerEvent", "<Action>,<arg1>,<arg2>")
                // Skip(1) skips past <Action> and passes any <args> that are present (separated by commas)
                switch (actionArgs[0])
                {
                    #region AZDOActions

                    case "GetWorkItemInfo":
                        Actions.AZDOActions.GetWorkItemInfo1(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetWorkItemInfo2":
                        Actions.AZDOActions.GetWorkItemInfo2(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetWorkItemRevisions":
                        Actions.AZDOActions.GetWorkItemRevisions(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "AddLinkedWorkItems":
                        Actions.AZDOActions.AddLinkedWorkItems1(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "AddLinkedWorkItems2":
                        Actions.AZDOActions.AddLinkedWorkItems2(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "AddLinkedWorkItemsExternal":
                        Actions.AZDOActions.AddLinkedWorkItemsExternal(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "QueryWorkItems":
                        Actions.AZDOActions.QueryWorkItems(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    #endregion

                    #region RoslynActions

                    case "CreateMethodShapes":
                        Actions.RoslynActions.CreateMethodShapes(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetClassInfo":
                        Actions.RoslynActions.GetClassInfo(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetProjectFileInfo":
                        Actions.RoslynActions.GetProjectFileInfo(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetSolutionFileInfo":
                        Actions.RoslynActions.GetSolutionFileInfo(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "GetSourceFileInfo":
                        Actions.RoslynActions.GetSourceFileInfo(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;


                    #endregion

                    #region Visio_Document Actions

                    case "CreatePluralSightCourseFileFromShape":
                        Actions.Visio_Document.CreatePluralSightCourseFileFromShape(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    #endregion

                    #region Visio_Page Actions

                    case "CreateActivityPage":
                        Actions.Visio_Page.CreateActivityPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateArtifactPage":
                        Actions.Visio_Page.CreateArtifactPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateDefaultLayersPage":
                        Actions.Visio_Page.CreateDefaultLayersPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateMetricPage":
                        Actions.Visio_Page.CreateMetricPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateNavigationLinksPage":
                        Actions.Visio_Page.CreateNavigationLinksPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreatePageBasePage":
                        Actions.Visio_Page.CreatePageBasePage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    // CreatePageForShape and LinkShapeToPage may be all we need unless special processing is needed.  
                    // Args can handle the common case of PreFix and Delimiter .e.g. L0-XYZ   Where L0 is Prefix and - is delimiter.
                    // Consider eliminating Create{ActivityPage,ArtifactPage,MetricPage,RolePage,ToolPage}

                    case "CreatePageForShape":
                        Actions.Visio_Page.CreatePageForShape(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateRolePage":
                        Actions.Visio_Page.CreateRolePage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "CreateToolPage":
                        Actions.Visio_Page.CreateToolPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "ToggleLayerLock":
                        Actions.Visio_Page.ToggleLayerLock(app, doc, page, shape, shapeu);
                        break;

                    case "ToggleLayerPrint":
                        Actions.Visio_Page.ToggleLayerPrint(app, doc, page, shape, shapeu);
                        break;

                    case "ToggleLayerVisibility":
                        Actions.Visio_Page.ToggleLayerVisibility(app, doc, page, shape, shapeu);
                        break;

                    case "UpdateGroups":
                        Actions.Visio_Page.UpdateGroupNameShapes(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "UpdateHasColorTags":
                        Actions.Visio_Page.UpdateHasColorTagsShapes(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "UpdateLayer":
                        Actions.Visio_Page.UpdateLayer(app, doc, page, shape, shapeu);
                        break;

                    #endregion

                    #region Visio_Shape Actions

                    case "LinkShapeToPage":
                        Actions.Visio_Shape.LinkShapeToPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "ListInvocationsInMethod":
                        Actions.Visio_Shape.ListInvocationsInMethod(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "ListMethodsInClass":
                        Actions.Visio_Shape.ListMethodsInClass(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    #endregion

                }
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
