﻿using SupportTools_Visio.Domain;
using VNC.Core.Mvvm;

namespace SupportTools_Visio.Presentation.ModelWrappers
{
    public class GeometryRowWrapper : ModelWrapper<Domain.GeometryRow>
    {
        public GeometryRowWrapper(GeometryRow model) : base(model)
        {
        }

        // TODO(crhodes)
        // This is going to take work
        public string Name { get { return GetValue<string>(); } set { SetValue(value); } }
        public string X { get { return GetValue<string>(); } set { SetValue(value); } }
        public string Y { get { return GetValue<string>(); } set { SetValue(value); } }
        public string A { get { return GetValue<string>(); } set { SetValue(value); } }
        public string B { get { return GetValue<string>(); } set { SetValue(value); } }
        public string C { get { return GetValue<string>(); } set { SetValue(value); } }
        public string D { get { return GetValue<string>(); } set { SetValue(value); } }
    }
}
