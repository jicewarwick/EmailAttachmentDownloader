# 邮箱附件下载工具
### Manual:
1. 在 `config.json` 里填写相关信息. 模板文件为 `config_example.json`
2. 运行 `attachment_archiever.py`.
3. 下载的附件保存在`config['storage_root']/sender/attachment_file_name`.  
若该文件已存在, 则改存`config['storage_root']/sender/{sent_time}_attachment_file_name`.

### Note:
- 只在QQ企业邮箱测试过. 其他邮箱还请自行测试, 需注意下载的附件文件名乱码问题

### Dependencies:
- tqdm

