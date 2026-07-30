# Web 端启动

开两个 PowerShell 窗口：

**后端**（先开）：

```powershell
cd d:\Coding\Project\PTS\zhongzi
.venv\Scripts\Activate.ps1
uvicorn web.backend.main:app --host 127.0.0.1 --port 8088 --reload
```

**前端**（后开）：

```powershell
cd d:\Coding\Project\PTS\zhongzi\web\frontend
npm run dev
```

然后浏览器访问 **http://localhost:3000**。
