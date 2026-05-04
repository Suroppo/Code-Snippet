using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 依赖注入的使用方式 {
    internal class 订单类 {
        private 商品接口 _商品;
        public 订单类(商品接口 商品) {
            Console.WriteLine("订单类");
            _商品 = 商品;
        }

        public IList<string> Show() {
            return [];
        }
    }

    interface 商品接口 {
        IList<string> 商品列表(string 分类, 用户接口 用户, 地址接口 地址);
    }

    internal class 商品类 : 商品接口 {

        public IList<string> 商品列表(string 分类, 用户接口 用户, 地址接口 地址) {
            Console.WriteLine("商品类");
            return [];
        }
    }

    interface 用户接口 {

    }

    internal class 用户类 : 用户接口 {
        private 等级接口 _等级;
        public 用户类(等级接口 等级) {
            Console.WriteLine("用户类");
            _等级 = 等级;
        }
    }

    interface 等级接口 {

    }

    internal class 等级类 : 等级接口 {
        public void 等级() {
            Console.WriteLine("等级类");
        }
    }

    interface 地址接口 {

    }

    internal class 地址类 : 地址接口 {
        public 地址类() {
            Console.WriteLine("地址类");
        }
    }
}
