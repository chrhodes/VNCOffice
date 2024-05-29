﻿using System;
using System.Windows;
using System.Windows.Input;
using Prism.Commands;
using SupportTools_Visio.Presentation.ModelWrappers;
using VNC;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ViewModels
{
    public class EditTextViewModel : ViewModelBase, IEditTextViewModel
    {
        public DelegateCommand DoSomething { get; private set; }
        public DelegateCommand<string> DoSomethingElse { get; private set; }
        public TextBlockFormatWrapper TextBlockFormat
        { 
            get => textBlockFormat; 
            set => textBlockFormat = value; 
        }
        //private string message = "Fox Lady";
        //public string Message
        //{ 
        //    get => message; 
        //    set => message = value; 
        //}

        public string Message { get; set; } = "Dr. Natalie Rene Rhodes";

        private TextBlockFormatWrapper textBlockFormat;

        public EditTextViewModel()
        {
            DoSomething = new DelegateCommand(OnDoSomethingExecute, OnDoSomethingCanExecute);
            DoSomethingElse = new DelegateCommand<string>(OnDoSomethingElseExecute, OnDoSomethingElseCanExecute);
            //DoSomething = new DelegateCommand(OnDoSomethingExecute);
            TextBlockFormat = new TextBlockFormatWrapper(new Domain.TextBlockFormat());
        }

        public void OnDoSomethingExecute()
        {
            Log.Trace("Enter", Common.LOG_CATEGORY);
            // Wrap a big, OMG, what have I done ???, undo around the whole thing !!!

            int undoScope = Globals.ThisAddIn.Application.BeginUndoScope("AddConnectionPoints");

            // Just need to pass in the model.

            Actions.Visio_Shape.UpdateTextSections(TextBlockFormat.Model);

            Globals.ThisAddIn.Application.EndUndoScope(undoScope, true);
            Log.Trace("Exit", Common.LOG_CATEGORY);
        }

        public Boolean OnDoSomethingCanExecute()
        {
            return true;
        }
        //public Action<string> OnDoSomethingElseExecute
        //{
        //    get
        //    {
        //        throw new NotImplementedException();
        //    }
        //}
        void OnDoSomethingElseExecute(string value)
        {
            switch (value)
            {
                case "0 pt":
                    SetMargins(value);
                    break;

                case "1 pt":
                    SetMargins(value);
                    break;

                case "2 pt":
                    SetMargins(value);
                    break;

                default:
                    MessageBox.Show($"Unknown tag: {value}");
                    break;
            }
        }

        bool OnDoSomethingElseCanExecute(string arg)
        {
            return true;
        }

        void SetMargins(string tag)
        {
            TextBlockFormat.TopMargin = tag;
            TextBlockFormat.BottomMargin = tag;
            TextBlockFormat.LeftMargin = tag;
            TextBlockFormat.RightMargin = tag;
        }
    }
}
