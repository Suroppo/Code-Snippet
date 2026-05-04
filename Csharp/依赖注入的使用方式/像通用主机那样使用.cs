using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;


namespace 依赖注入的使用方式 {
    internal class 通用主机式 {
        internal static void 使用() {
            var builder = AppBuilder.CreateBuilder();
            builder.ConfigureService(services => {
                // 注册所有的服务及实现, 并确定其生命周期
                services.AddTransient<订单类>();
                services.AddTransient<商品接口, 商品类>();
                services.AddTransient<用户接口, 用户类>();
                services.AddTransient<等级接口, 等级类>();
                services.AddTransient<地址接口, 地址类>();
            });
            // 通过扩展类注册服务及实现, 可以实现把相关的接口或类放在一起注册
            builder.添加定时任务服务();

            builder.Start();
        }
    }

    public interface IServiceBuilder { 
        IServiceBuilder ConfigureService(Action<IServiceCollection> services);
    }

    class AppBuilder : IServiceBuilder {
        private IServiceCollection _services;
        public IServiceBuilder ConfigureService(Action<IServiceCollection> services) {
            return this;
        }

        public IServiceBuilder ConfigureService(Action<IConfiguration, IServiceCollection> services) {
            return this;
        }

        private AppBuilder() {
            _services = new ServiceCollection();
        }
        public static AppBuilder CreateBuilder() {
            // 这里可以添加一些 应用运行时的相关逻辑
            // 比如构建全局对象
            //  注入默认的服务, 日志, 配置等
            return new AppBuilder();
        }

        public void Start() {
            var provider = _services.BuildServiceProvider();

            using (var rootScope = provider.CreateScope()) {
                var app = rootScope.ServiceProvider.GetRequiredService<订单类>();
                app.Show();

            }
        }
    }

    public static class ServiceBuilderExtensions {
        // 模拟分包注册
        public static IServiceBuilder 添加定时任务服务(this IServiceBuilder builder) {
            builder.ConfigureService(services =>
            {
                // services.AddTransient<接口, 类>();
                // 注册其他业务服务...
            });
            return builder;
        }
    }




}
