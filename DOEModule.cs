using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Services;
using MaxChemical.Modules.DOE.ViewModels;
using MaxChemical.Modules.DOE.Views;
using Prism.Ioc;
using Prism.Modularity;


namespace MaxChemical.Modules.DOE
{
    /// <summary>
    /// DOE 模块入口
    /// 
    ///  修改: 新增 IMultiResponseGPRService 注册
    /// </summary>
    public class DOEModule : IModule
    {
        public void OnInitialized(IContainerProvider containerProvider) { }

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            // 数据仓储
            containerRegistry.RegisterSingleton<IDOERepository, DOERepository>();

            // 核心服务
            containerRegistry.RegisterSingleton<IDOEDesignService, DOEDesignService>();
            containerRegistry.RegisterSingleton<IDOEExecutionService, DOEBatchExecutor>();
            containerRegistry.RegisterSingleton<IGPRModelService, GPRModelService>();
            containerRegistry.RegisterSingleton<IMultiResponseGPRService, MultiResponseGPRService>();  //  新增
            containerRegistry.RegisterSingleton<IDOEAnalysisService, DOEAnalysisService>();
            containerRegistry.RegisterSingleton<DOEExportService>();
            containerRegistry.RegisterSingleton<IDesirabilityService, DesirabilityService>();
            containerRegistry.RegisterSingleton<IModelRouter, ModelRouter>();
            // ═══ 项目层服务 ═══
            ProjectServiceRegistration.RegisterProjectServices(containerRegistry);
            // ViewModel
            containerRegistry.Register<DOEMainViewModel>();
            containerRegistry.Register<DOEOverviewViewModel>();
            containerRegistry.Register<DOEDesignWizardViewModel>();
            containerRegistry.Register<DOEExecutionDashboardViewModel>();
            containerRegistry.Register<DOEModelAnalysisViewModel>();
            containerRegistry.Register<DOEHistoryViewModel>();
            containerRegistry.Register<DOEProjectDetailViewModel>();
            
            // View
            containerRegistry.Register<DOEMainView>();
            containerRegistry.Register<DOEOverviewView>();
            containerRegistry.Register<DOEDesignWizardView>();
            containerRegistry.Register<DOEExecutionDashboardView>();
            containerRegistry.Register<DOEModelAnalysisView>();
            containerRegistry.Register<DOEHistoryView>();
            containerRegistry.Register<DOEProjectDetailDialog>();
        }
    }
}