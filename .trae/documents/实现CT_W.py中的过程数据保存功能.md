1. **在OSAController类中添加保存功能**

   * 添加`_query`方法：带重试的查询功能，提高稳定性

   * 添加`_opc_wait`方法：等待操作完成

   * 添加`save_screenshot`方法：保存当前截图

   * 添加`save_curve_atrace`方法：使用":MMEMory:STORe:ATRace"指令保存所有曲线

2. **在TestRunner类中添加辅助方法**

   * 添加`_ensure_process_data_dir`方法：确保"过程数据"文件夹存在

   * 添加`_save_process_data`方法：封装保存截图和曲线的逻辑

3. **修改测试运行逻辑**

   * 在`run_group1`方法的每次温度步进后，调用`_save_process_data`

   * 在`run_group2`方法的每次电流步进后，调用`_save_process_data`

   * 使用包含测试组、温度/电流、时间戳的文件名格式

4. **文件命名规范**

   * 截图命名：`GroupX_TempXXX_CurrXXX_YYYYMMDD_HHMMSS.bmp`

   * 曲线命名：`GroupX_TempXXX_CurrXXX_YYYYMMDD_HHMMSS.csv`

   * 其中X表示测试组，XXX表示温度或电流值

5. **保存路径**

   * 在主保存路径下创建"过程数据"子文件夹

   * 所有过程数据保存到该子文件夹中

6. **异常处理**

   * 为保存操作添加适当的异常处理

   * 确保保存失败不会影响主测试流程

   * 记录保存操作的日志信息

