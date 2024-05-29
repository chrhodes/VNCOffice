﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class ChangeShapeBehaviorWrapper : ModelWrapper<ChangeShapeBehavior>
    {
        public ChangeShapeBehaviorWrapper()
        {
        }
        public ChangeShapeBehaviorWrapper(ChangeShapeBehavior model) : base(model)
        {
        }

        public string ReplaceLockShapeData { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockText { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceLockFormat { get { return GetValue<string>(); } set { SetValue(value); } }
        public string ReplaceCopyCells { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
