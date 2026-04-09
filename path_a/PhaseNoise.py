from pywinauto import Desktop
import time

def click_button():
    try:
        desktop = Desktop(backend="win32")
        # 绑定目标窗口
        matched_windows = desktop.windows(
            class_name="SunAwtFrame", 
            title="LaserNoiseMeasurement_Manufactory",
            visible_only=True
        )
        if not matched_windows:
            print("未找到目标窗口。")
            return False
            
        win = desktop.window(handle=matched_windows[0].handle)
        
        # 激活窗口
        try:
            win.set_focus()
        except Exception:
            pass
        time.sleep(0.5) 
        
        # 使用算好的相对坐标点击
        win.click_input(coords=(296, 675))
        print("已成功点击目标按钮！(相对坐标: X=296, Y=675)")
        return True
        
    except Exception as e:
        print(f"点击失败: {e}")
        return False

# 执行点击
click_button()