# SkyGuard 数据字典

本数据字典区分真实来源与演示样本。

## national
- 行数：84
- 字段：year, 指标编码, 指标名称, value, unit, previous_year_value, yoy_pct, 数据性质, platform_relevance, 来源索引, 来源链接, note
- data_nature取值：official_statistic

## regional
- 行数：42
- 字段：year, region, category_code, category_name_cn, value, unit, national_total, share_pct, previous_year_value, yoy_abs, yoy_pct, 数据性质, 来源索引, 来源链接, note
- data_nature取值：official_statistic

## airport
- 行数：81
- 字段：year, dimension, region_or_bucket, indicator, value, unit, 来源索引, 来源链接, 数据性质, platform_relevance, share_pct, share_meaning
- data_nature取值：official_statistic

## cloud
- 行数：39
- 字段：year, dimension, indicator, value, unit, 数据性质, 来源索引, 来源链接, note, platform_relevance
- data_nature取值：official_statistic, official_text, derived

## city
- 行数：158
- 字段：entity, entity_level, policy_or_case_doc, target_year, category, indicator, comparator, value, unit, 数据性质, platform_relevance, 来源索引, 来源链接, note
- data_nature取值：policy_target, policy_measure, public_research_citation, public_case_statistic, public_case_event, policy_task

## plans
- 行数：8000
- 字段：flight_plan_id, operator_id, aircraft_id, aircraft_type, city, city_code, region, route_id, scenario, apply_time, planned_takeoff_time, planned_landing_time, planned_duration_min, planned_altitude_m, planned_distance_km, night_flight, payload_kg, approval_status, preflight_risk_score, preflight_risk_level, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## telemetry
- 行数：30000
- 字段：telemetry_id, flight_plan_id, aircraft_id, city, city_code, route_id, timestamp, lat, lon, altitude_m, speed_mps, heading_deg, battery_pct, communication_quality, gnss_satellite_count, remote_id_valid, geofence_status, data_quality_flag, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## alerts
- 行数：6000
- 字段：alert_id, flight_plan_id, aircraft_id, city, city_code, route_id, timestamp, alert_type, severity, lat, lon, altitude_m, detection_source, confidence, disposal_status, response_time_min, closed_loop, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## risk
- 行数：6000
- 字段：sample_id, grid_id, city, city_code, timestamp_hour, airspace_class, population_density_index, flight_density_index, sensor_coverage_score, communication_quality_score, wind_speed_mps, rainfall_mm_h, visibility_km, historical_alert_count_24h, computed_risk_score, risk_event_label, 数据性质, source_basis
- data_nature取值：simulated_model_training_sample

## grid
- 行数：10000
- 字段：grid_id, city, city_code, region, centroid_lat, centroid_lon, airspace_class, max_allowed_altitude_m, nearby_airport_distance_km, population_density_index, flight_density_index, sensor_coverage_score, communication_quality_score, weather_risk_score, computed_risk_score, computed_risk_level, no_fly_or_control_reason, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## sensors
- 行数：1000
- 字段：station_id, city, city_code, region, sensor_type, lat, lon, coverage_radius_km, height_m, online_status, health_score, last_maintenance_date, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## vertiports
- 行数：900
- 字段：vertiport_id, city, city_code, region, vertiport_type, lat, lon, pad_count, charging_available, daily_capacity_sorties, ground_access_level, approval_status, 数据性质, source_basis
- data_nature取值：simulated_demo_sample

## routes
- 行数：1200
- 字段：route_id, city, city_code, region, scenario, origin_vertiport_id, destination_vertiport_id, route_distance_km, planned_altitude_m, corridor_width_m, daily_planned_sorties, route_status, risk_control_requirement, 数据性质, source_basis
- data_nature取值：simulated_demo_sample