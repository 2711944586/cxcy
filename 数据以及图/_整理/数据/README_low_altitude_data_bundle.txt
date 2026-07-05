低空经济城市空域安全感知平台 - 已验证CSV数据包

本数据包用于课堂汇报、数据分析、报表演示或低空空域安全感知平台原型。数据均来自公开来源，未人工伪造。

文件列表：

1. opensky_flightlist_202001.csv.gz
   来源：OpenSky Network / ActiveConclusion COVID19_AirTraffic 项目
   格式：压缩CSV
   本地实测行数：1,519,103 行数据 + 1 行表头
   字段：callsign, number, icao24, registration, typecode, origin, destination, firstseen, lastseen, day
   用途：飞行活动统计、空域态势、时序报表、起终点关联、航班密度分析。
   原始链接：https://raw.githubusercontent.com/ActiveConclusion/COVID19_AirTraffic/master/opensky_data/flightlist_20200101_20200131.csv.gz

2. ourairports_airports.csv
   来源：OurAirports Open Data
   本地实测行数：85,644 行数据 + 1 行表头
   用途：机场/起降设施基础数据、空间约束、起降点选址分析。
   原始链接：https://davidmegginson.github.io/ourairports-data/airports.csv

3. ourairports_runways.csv
   来源：OurAirports Open Data
   本地实测行数：48,059 行数据 + 1 行表头
   用途：跑道/起降能力、机场容量、航线安全约束分析。
   原始链接：https://davidmegginson.github.io/ourairports-data/runways.csv

4. ourairports_navaids.csv
   来源：OurAirports Open Data
   本地实测行数：11,009 行数据 + 1 行表头
   用途：导航台分布、航空基础设施地图、航线规划辅助。
   原始链接：https://davidmegginson.github.io/ourairports-data/navaids.csv

建议组合：
OpenSky flightlist 作为“飞行运行主表”，OurAirports 作为“基础设施维表”，再叠加 Open-Meteo 或 Meteostat 的小时级天气CSV，就能形成一套比较完整的低空/空域安全感知分析数据底座。
