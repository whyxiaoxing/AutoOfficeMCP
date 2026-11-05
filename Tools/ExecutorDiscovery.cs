using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Office.Contracts;

namespace AutoOfficeMCP.Tools
{
    /// <summary>
    /// 通过反射发现 Office / Xlsx 程序集中的 IToolExecutor 实现，并为工具描述符路由执行器。
    /// 约定：不硬编码任何 T03/T04 执行器类名；无执行器认领时回退 StubToolExecutor。
    /// </summary>
    public sealed class ExecutorDiscovery
    {
        private static readonly string[] ExecutorAssemblyNames = new string[] { "Office", "Xlsx", "PowerPointe" };

        private readonly List<IToolExecutor> _executors;
        private readonly List<IAsyncToolExecutor> _asyncExecutors;
        private readonly StubToolExecutor _fallback;

        public ExecutorDiscovery()
        {
            _executors = new List<IToolExecutor>();
            _asyncExecutors = new List<IAsyncToolExecutor>();
            _fallback = new StubToolExecutor();

            foreach (string assemblyName in ExecutorAssemblyNames)
            {
                LoadFromAssembly(assemblyName);
            }

            // T06-Common 执行器位于主程序集，必须与 Office/Xlsx 一并扫描。
            LoadFromAssembly(Assembly.GetExecutingAssembly(), "AutoOfficeMCP（主程序集）");
        }

        public IReadOnlyList<IToolExecutor> Executors
        {
            get { return _executors; }
        }

        /// <summary>
        /// 按 descriptor.appType 与 CanExecute 找到执行器；找不到时回退 stub。
        /// 执行器异常在这里收敛为 ToolResult.Fail，避免把异常抛回协议层。
        /// </summary>
        public ToolResult Execute(ToolDescriptor descriptor, JObject arguments)
        {
            if (descriptor == null)
            {
                return ToolResult.Fail("INVALID_ARGUMENT", "执行器路由失败：工具描述符为空。", false);
            }

            IToolExecutor executor = FindExecutor(descriptor.ToolName, descriptor.AppType);
            if (executor == null)
            {
                return _fallback.Execute(descriptor.ToolName, arguments);
            }

            try
            {
                ToolResult result = executor.Execute(descriptor.ToolName, arguments);
                if (result == null)
                {
                    return ToolResult.Fail("INTERNAL_ERROR", "执行器返回了空结果：" + descriptor.ToolName, true);
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("执行器抛出未捕获异常：" + executor.GetType().FullName
                                        + " / " + descriptor.ToolName + Environment.NewLine + ex);
                ToolResult failure = ToolResult.Fail("INTERNAL_ERROR", "工具执行异常：" + ex.Message, true);
                failure.Diagnostic = Office.OfficeErrorMapper.FormatExceptionChain(ex);
                return failure;
            }
        }

        /// <summary>是否存在异步执行器认领该工具；用于协议层选择 await 路径。</summary>
        public bool HasAsyncExecutor(ToolDescriptor descriptor)
        {
            if (descriptor == null)
            {
                return false;
            }
            return FindAsyncExecutor(descriptor.ToolName, descriptor.AppType) != null;
        }

        /// <summary>
        /// 是否纯本地工具（无 COM、无需显式 file）。用于固定进程 worker 模式绕过 worker 投递。
        /// 执行器未实现 ILocalToolExecutor 时返回 false。
        /// </summary>
        public bool IsLocalTool(ToolDescriptor descriptor, JObject arguments)
        {
            if (descriptor == null)
            {
                return false;
            }
            IToolExecutor executor = FindExecutor(descriptor.ToolName, descriptor.AppType);
            ILocalToolExecutor local = executor as ILocalToolExecutor;
            if (local == null)
            {
                return false;
            }
            try
            {
                return local.IsLocalTool(descriptor.ToolName, arguments);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("执行器 IsLocalTool 异常：" + executor.GetType().FullName + " / "
                                        + descriptor.ToolName + "；" + ex.Message);
                return false;
            }
        }

        /// <summary>异步执行入口：优先 async 执行器；没有时回退同步 Execute。</summary>
        public async Task<ToolResult> ExecuteAsync(ToolDescriptor descriptor, JObject arguments)
        {
            if (descriptor == null)
            {
                return ToolResult.Fail("INVALID_ARGUMENT", "执行器路由失败：工具描述符为空。", false);
            }

            IAsyncToolExecutor executor = FindAsyncExecutor(descriptor.ToolName, descriptor.AppType);
            if (executor == null)
            {
                return Execute(descriptor, arguments);
            }

            try
            {
                ToolResult result = await executor.ExecuteAsync(descriptor.ToolName, arguments).ConfigureAwait(false);
                if (result == null)
                {
                    return ToolResult.Fail("INTERNAL_ERROR", "异步执行器返回了空结果：" + descriptor.ToolName, true);
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("异步执行器抛出未捕获异常：" + executor.GetType().FullName
                                        + " / " + descriptor.ToolName + Environment.NewLine + ex);
                ToolResult failure = ToolResult.Fail("INTERNAL_ERROR", "工具执行异常：" + ex.Message, true);
                failure.Diagnostic = Office.OfficeErrorMapper.FormatExceptionChain(ex);
                return failure;
            }
        }

        private IAsyncToolExecutor FindAsyncExecutor(string toolName, string appType)
        {
            foreach (IAsyncToolExecutor executor in _asyncExecutors)
            {
                if (!string.Equals(executor.AppType, appType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                try
                {
                    if (executor.CanExecute(toolName))
                    {
                        return executor;
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("异步执行器 CanExecute 异常：" + executor.GetType().FullName +
                                            " / " + toolName + "；" + ex.Message);
                }
            }
            return null;
        }

        private IToolExecutor FindExecutor(string toolName, string appType)
        {
            foreach (IToolExecutor executor in _executors)
            {
                if (!string.Equals(executor.AppType, appType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (SafeCanExecute(executor, toolName))
                {
                    return executor;
                }
            }
            return null;
        }

        private static bool SafeCanExecute(IToolExecutor executor, string toolName)
        {
            try
            {
                return executor.CanExecute(toolName);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("执行器 CanExecute 异常：" + executor.GetType().FullName +
                                        " / " + toolName + "；" + ex.Message);
                return false;
            }
        }

        private void LoadFromAssembly(string assemblyName)
        {
            Assembly assembly;
            try
            {
                assembly = Assembly.Load(new AssemblyName(assemblyName));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("无法加载执行器程序集 " + assemblyName + "：" + ex.Message);
                return;
            }

            LoadFromAssembly(assembly, assemblyName);
        }

        private void LoadFromAssembly(Assembly assembly, string assemblyName)
        {
            if (assembly == null)
            {
                return;
            }

            Type[] types;
            try
            {
                types = assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                List<Type> loadedTypes = new List<Type>();
                if (ex.Types != null)
                {
                    foreach (Type candidate in ex.Types)
                    {
                        if (candidate != null)
                        {
                            loadedTypes.Add(candidate);
                        }
                    }
                }
                types = loadedTypes.ToArray();
                Console.Error.WriteLine("程序集 " + assemblyName + " 部分类型加载失败（可能缺少 COM 依赖），"
                                        + "仅扫描成功加载的 " + types.Length + " 个类型。");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("无法枚举程序集类型 " + assemblyName + "：" + ex.Message);
                return;
            }

            foreach (Type type in types)
            {
                if (type == null || !type.IsClass || type.IsAbstract)
                {
                    continue;
                }
                if (!typeof(IToolExecutor).IsAssignableFrom(type))
                {
                    continue;
                }
                if (type == typeof(StubToolExecutor))
                {
                    continue;
                }

                ConstructorInfo constructor = type.GetConstructor(Type.EmptyTypes);
                if (constructor == null)
                {
                    Console.Error.WriteLine("跳过缺少无参构造的 IToolExecutor：" + type.FullName);
                    continue;
                }

                try
                {
                    IToolExecutor executor = Activator.CreateInstance(type) as IToolExecutor;
                    if (executor == null)
                    {
                        Console.Error.WriteLine("实例化后无法转换为 IToolExecutor：" + type.FullName);
                        continue;
                    }
                    _executors.Add(executor);
                    IAsyncToolExecutor asyncExecutor = executor as IAsyncToolExecutor;
                    if (asyncExecutor != null)
                    {
                        _asyncExecutors.Add(asyncExecutor);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("实例化执行器失败：" + type.FullName +
                                            Environment.NewLine + ex);
                }
            }
        }
    }
}
