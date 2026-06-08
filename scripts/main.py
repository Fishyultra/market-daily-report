import os
import io
import requests
import pandas as pd
import matplotlib.pyplot as plt
import cloudinary
import cloudinary.uploader
from datetime import datetime

# ==================== 配置区 ====================
PUSHPLUS_TOKEN = os.environ.get("PUSHPLUS_TOKEN")

cloudinary.config(
    cloud_name=os.environ.get("CLOUDINARY_CLOUD_NAME"),
    api_key=os.environ.get("CLOUDINARY_API_KEY"),
    api_secret=os.environ.get("CLOUDINARY_API_SECRET"),
    secure=True
)

TODAY = datetime.now().strftime("%Y-%m-%d")

# ==================== 工具函数 ====================
def upload_image_to_cloudinary(image_buffer, public_id):
    try:
        image_buffer.seek(0)
        result = cloudinary.uploader.upload(
            image_buffer,
            public_id=public_id,
            overwrite=True,
            resource_type="image"
        )
        return result.get("secure_url")
    except Exception as e:
        print(f"Cloudinary 上传失败: {e}")
        return None

def pushplus_send(title, content, token):
    url = "http://www.pushplus.plus/send"
    data = {
        "token": token,
        "title": title,
        "content": content,
        "template": "markdown"
    }
    try:
        resp = requests.post(url, data=data, timeout=30)
        print(f"Pushplus 响应: {resp.text}")
        return resp.json()
    except Exception as e:
        print(f"Pushplus 推送失败: {e}")
        return None

# ==================== 主逻辑 ====================
def main():
    print(f"开始读取 Excel 并生成图表...")
    
    # 读取仓库根目录下的 data.xlsx
    try:
        df = pd.read_excel("data.xlsx", engine="openpyxl")
        print(f"成功读取数据，共 {len(df)} 行，列：{list(df.columns)}")
    except Exception as e:
        print(f"读取 Excel 失败: {e}")
        return
    
    if len(df.columns) < 2:
        print("Excel 至少需要 2 列数据（第一列X轴，后面列Y轴）")
        return
    
    # 作图：第一列作为X轴，其余列作为Y轴
    x_col = df.columns[0]
    y_cols = df.columns[1:]
    
    fig, ax = plt.subplots(figsize=(10, 5))
    colors = ["#E74C3C", "#3498DB", "#2ECC71", "#F39C12", "#9B59B6"]
    
    for i, col in enumerate(y_cols):
        ax.plot(df[x_col], df[col], label=str(col), color=colors[i % len(colors)], 
                marker='o', linewidth=2, markersize=4)
    
    ax.set_title(f"Excel Data Chart ({TODAY})", fontsize=14, fontweight='bold')
    ax.set_xlabel(str(x_col))
    ax.legend(loc='best')
    ax.grid(True, alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    # 保存到内存
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close()
    
    # 上传 Cloudinary
    chart_url = upload_image_to_cloudinary(buf, f"excel_chart_{TODAY}")
    if not chart_url:
        print("图表上传失败，终止")
        return
    
    print(f"图表 URL: {chart_url}")
    
    # 组装推送内容
    markdown_content = f"""## 📊 Excel Chart Test ({TODAY})

![Chart]({chart_url})

Data source: data.xlsx ({len(df)} rows, {len(df.columns)} columns)
"""
    
    # 推送
    title = f"Excel Chart Test {TODAY}"
    result = pushplus_send(title, markdown_content, PUSHPLUS_TOKEN)
    
    if result and result.get("code") == 200:
        print("推送成功！")
    else:
        print(f"推送失败: {result}")

if __name__ == "__main__":
    main()
