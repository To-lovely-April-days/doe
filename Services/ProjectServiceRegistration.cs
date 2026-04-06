using Prism.Ioc;
using MaxChemical.Modules.DOE.Services;

namespace MaxChemical.Modules.DOE
{
    /// <summary>
    /// ★ 新增: 项目层服务注册
    /// 
    /// 在 DOEModule.RegisterTypes() 中添加以下注册:
    /// 
    /// <code>
    /// // ═══ 项目层服务 ═══
    /// containerRegistry.RegisterSingleton&lt;IProjectDecisionEngine, ProjectDecisionEngine&gt;();
    /// </code>
    /// 
    /// 不需要注册 Repository（已有的 DOERepository 通过 partial class 自动包含项目方法）。
    /// 不需要注册 DOEDesignService（通过 partial class 自动包含新方法）。
    /// 
    /// 完整的注册示例:
    /// </summary>
    public static class ProjectServiceRegistration
    {
        /// <summary>
        /// 在 DOEModule.RegisterTypes 中调用此方法注册项目层服务
        /// </summary>
        public static void RegisterProjectServices(IContainerRegistry containerRegistry)
        {
            // 项目决策引擎（单例）
            containerRegistry.RegisterSingleton<IProjectDecisionEngine, ProjectDecisionEngine>();
        }
    }
}
