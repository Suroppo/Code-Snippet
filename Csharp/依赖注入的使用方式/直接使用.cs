using Microsoft.Extensions.DependencyInjection;


namespace 依赖注入的使用方式 {
    internal class 直接式 {
        internal static void 使用() {
            var services = new ServiceCollection();
            // 注册所有的服务及实现, 并确定其生命周期
            services.AddTransient<订单类>();
            services.AddTransient<商品接口, 商品类>();
            services.AddTransient<用户接口, 用户类>();
            services.AddTransient<等级接口, 等级类>();
            services.AddTransient<地址接口, 地址类>();

            // 构建服务的提供者
            var serviceProvider = services.BuildServiceProvider();

            // 
            using (var rootSer = serviceProvider.CreateScope()) {
                // 这里只用获取这个入口类的实例就可以了, 在该类后继的依赖实例, 依赖注入将自动创建其实例
                rootSer.ServiceProvider.GetRequiredService<订单类>();
            }
        }
    }



}
