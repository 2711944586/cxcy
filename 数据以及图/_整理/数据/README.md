# 低空经济城市空域安全感知平台资料包

整理日期：2026-06-28

## 文件夹内容

- `images/`：真实平台/低空运行相关图片 4 张，含 `图片来源标注.csv`。
- `government_notice/`：中国政府网《无人驾驶航空器飞行管理暂行条例》HTML 原文和公告摘要图。
- `data/`：从 Kaggle 下载的 UAV/低空安全感知参考 CSV 数据集，含压缩包、解压 CSV 和 `Kaggle数据集清单.csv`。
- `metadata/`：Kaggle API 返回的检索和详情 JSON，便于追溯。

## 政府公告

- 标题：无人驾驶航空器飞行管理暂行条例
- 发文机关：国务院、中央军委
- 发文字号：国令第761号
- 成文日期：2023年05月31日
- 发布日期：2023年06月28日
- 来源：中国政府网
- 原文：https://www.gov.cn/zhengce/content/202306/content_6888799.htm

## Kaggle CSV 数据说明

已下载并解压 3 个 Kaggle 数据集，共 4 个 CSV：

1. UAV Communication Monitoring Dataset：通信强度、频率、带宽、调制方式、加密协议、延迟、标签等字段，可用于低空通信链路安全感知。
2. UAV Autonomous Navigation Dataset：经纬度、高度、IMU、激光雷达距离、速度、风速、电量、障碍物检测等字段，可用于低空运行态势感知。
3. Cyber Network Intrusion Dataset for UAVs：无人机网络/Wi-Fi 流量字段和攻击分类标签，可用于平台网络安全告警和异常检测。

详见 `data/Kaggle数据集清单.csv`。

## 图片说明

图片均保留为下载到的原始 JPG/PNG 文件，未加水印、未二次压缩。来源网站、来源页面和图片直链见 `images/图片来源标注.csv`。

## 注意

这些资料用于课程/方案调研和非商业引用整理。正式公开发布或商业使用前，请再次核对各来源网站的版权、转载和授权条款。
