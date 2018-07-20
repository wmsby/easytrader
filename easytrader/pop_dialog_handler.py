# coding:utf-8
import re
import time

from . import exceptions


class PopDialogHandler:
    def __init__(self, app):
        self._app = app

    def handle(self, title):
        if any(s in title for s in {"提示信息", "委托确认", "网上交易用户协议"}):
            self._submit_by_shortcut()

        elif "提示" in title:
            content = self._extract_content()
            self._submit_by_click()
            return {"message": content}

        else:
            content = self._extract_content()
            self._close()
            return {"message": "unknown message: {}".format(content)}

    def _extract_content(self):
        for c in range(20):
            sss = time.time()
            try:
                test = self._app.top_window()
                test.wait("exists visible enabled", 0.05)
                return test.Static.window_text()
            except Exception as e:
                print('_extract_content', e)
            zzz = time.time()
            if (zzz-sss) < 0.05:
                time.sleep(0.05-(zzz-sss))
                
        return test.Static.window_text()

    def _extract_entrust_id(self, content):
        return re.search(r"\d+", content).group()

    def _submit_by_click(self):
        self._app.top_window()["确定"].click()

    def _submit_by_shortcut(self):
        self._app.top_window().type_keys("%Y")

    def _submit_by_shortcut_yes(self):  # 点击 是
        self._app.top_window().type_keys("%Y")

    def _submit_by_shortcut_no(self):   # 点击 否
        self._app.top_window().type_keys("%N")
        
    def _close(self):
        self._app.top_window().close()


class TradePopDialogHandler(PopDialogHandler):
    def handle(self, title):
        if title == "委托确认":
            self._submit_by_shortcut_yes()
            time.sleep(0.1)

        elif title == "提示信息":
            content = self._extract_content()
            if "超出涨跌停" in content:
                self._submit_by_shortcut_no()
                return {"failure": content}
            elif "委托价格的小数部分应为" in content:
                self._submit_by_shortcut_no()
                return {"failure": content}
            else:
                self._submit_by_shortcut_yes()
                time.sleep(0.1)

        elif title == "提示":
            content = self._extract_content()
            if "成功" in content:
                entrust_no = self._extract_entrust_id(content)
                self._submit_by_click()
                return {"success": entrust_no}
            else:
                self._submit_by_click()
                return {"failure": content}
        else:
            self._close()
