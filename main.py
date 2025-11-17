# main.py
from web_api import app # 从你的 Web API 文件中导入那个 app 实例
import os

if __name__ == "__main__":
    """
    这是新的程序入口。
    它只负责一件事：启动 Flask 服务器。
    """
    print("--- 启动碳报告 Web 服务 ---")
    print("你现在可以通过浏览器访问了。")

    # 告诉 Flask 在哪个地址和端口上"监听"
    # host='0.0.0.0' 意味着局域网内的其他人也能访问
    # debug=True 意味着你修改代码后服务器会自动重启，很方便
    port = int(os.getenv('PORT', 5071))
    app.run(host='0.0.0.0', port=port, debug=True)