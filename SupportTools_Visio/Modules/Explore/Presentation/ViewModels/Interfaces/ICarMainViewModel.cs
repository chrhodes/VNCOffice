﻿using System.Threading.Tasks;

using VNC.Core.Mvvm;

namespace Explore.Presentation.ViewModels
{
    public interface ICarMainViewModel : IViewModel
    {
        Task LoadAsync();
    }
}