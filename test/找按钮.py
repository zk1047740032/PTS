from pywinauto.application import Application
app = Application(backend="uia").connect(title_re=r".*WLM LongTerm graph.*")
win = app.window(title_re=r".*WLM LongTerm graph.*")
charts = win.descendants(class_name="TChart")
for i, c in enumerate(charts):
    try:
        print(f"found_index={i} | 坐标: {c.rectangle()}")
    except:
        pass