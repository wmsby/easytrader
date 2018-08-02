# -*- coding: utf-8 -*-
import abc
import functools
import os
import sys
import time
import pandas as pd
import easyutils
import pywinauto
import win32gui, win32com.client

from . import grid_data_get_strategy
from . import helpers
from . import pop_dialog_handler
from .config import client
from .log import log

if not sys.platform.startswith("darwin"):
    import pywinauto
    import pywinauto.clipboard


class IClientTrader(abc.ABC):
    @property
    @abc.abstractmethod
    def app(self):
        """Return current app instance"""
        pass

    @property
    @abc.abstractmethod
    def main(self):
        """Return current main window instance"""
        pass

    @property
    @abc.abstractmethod
    def config(self):
        """Return current config instance"""
        pass

    @abc.abstractmethod
    def wait(self, seconds: int):
        """Wait for operation return"""
        pass

    @property
    @abc.abstractmethod
    def grid_data_get_strategy(self):
        """
        :return: Implement class of IGridDataGetStrategy
        :rtype: grid_data.get_strategy.IGridDataGetStrategy
        """
        pass

    @grid_data_get_strategy.setter
    @abc.abstractmethod
    def grid_data_get_strategy(self, strategy_cls):
        """
        :param strategy_cls: Grid data get strategy
        :type strategy_cls: grid_data.get_strategy.IGridDataGetStrategy
        :return: formatted grid data
        :rtype: list[dict]
        """
        pass


class ClientTrader(IClientTrader):
    def __init__(self):
        self._config = client.create(self.broker_type)
        self._app = None
        self._main = None
        self._main_handle = None
        self._left_treeview = None
        self._pwindow = None
        self.grid_data_get_strategy = grid_data_get_strategy.CopyStrategy

    @property
    def app(self):
        return self._app

    @property
    def main(self):
        return self._main

    @property
    def config(self):
        return self._config

    @property
    def grid_data_get_strategy(self):
        return self._grid_data_get_strategy

    @grid_data_get_strategy.setter
    def grid_data_get_strategy(self, strategy_cls):
        if not issubclass(
            strategy_cls, grid_data_get_strategy.IGridDataGetStrategy
        ):
            raise TypeError(
                "Strategy must be implement class of IGridDataGetStrategy"
            )
        self._grid_data_get_strategy = strategy_cls(self)

    def connect(self, exe_path=None, **kwargs):
        """
        直接连接登陆后的客户端
        :param exe_path: 客户端路径类似 r'C:\\htzqzyb2\\xiadan.exe', 默认 r'C:\\htzqzyb2\\xiadan.exe'
        :return:
        """
        connect_path = exe_path or self._config.DEFAULT_EXE_PATH
        if connect_path is None:
            raise ValueError(
                "参数 exe_path 未设置，请设置客户端对应的 exe 地址,类似 C:\\客户端安装目录\\xiadan.exe"
            )

        self._app = pywinauto.Application().connect(
            path=connect_path, timeout=10
        )
        self._close_prompt_windows()

        self._main = self._app.window_(title_re="网上股票交易系统")
        self._main.wait('exists enabled visible ready')
        
        self._main_handle = self._main.handle
        
        self._left_treeview = self._main.window_(control_id=129, class_name="SysTreeView32") 
        self._left_treeview.wait('exists enabled visible ready')
        
        self._pwindow = self._main.window(control_id=59649, class_name='#32770')
        self._pwindow.wait('exists enabled visible ready')
        
    # check top_window
    def _check_top_window(self):
        """需要6ms"""
        c = 0
        test_0 = self._main.wrapper_object()
        test_1 = self._app.top_window().wrapper_object()
        while c < 5 and test_1 != test_0:
            c += 1
            test_1.close()
            test_1 = self._app.top_window().wrapper_object()
            
    def _close_prompt_windows(self):
        """功能同_check_top_window, 需要2ms, 不太可靠"""
        for w in self._app.windows(class_name="#32770"):
            if "网上交易系统" not in w.window_text():
                w.close()
        
    @property
    def broker_type(self):
        return "ths"

    @functools.lru_cache()
    def _get_left_treeview_ready(self):
        for c in range(2):
            try:
                self._left_treeview.wait("ready", 1)
                break
            except:
                log.warning('_left_treeview.wait Exception')
                self._bring_main_foreground()
                self._check_top_window()
            
    def _switch_left_menus(self, path):
        def left_menus_check():
            try:
                if self._left_treeview.IsSelected(path):
                    return True
                else:
                    return False
            except Exception as e:
                log.warning('_switch_left_menus: {}'.format(e))
                self._get_left_treeview_ready()
                return False
            
        test = ''.join(path)
        for c in range(2):
            if 'F1' in test:
                self._main.TypeKeys("{F1}")
                if left_menus_check():
                    break
            elif 'F2' in test:
                self._main.TypeKeys("{F2}")
                if left_menus_check():
                    break
            elif 'F3' in test:
                self._main.TypeKeys("{F3}")
                if left_menus_check():
                    break
            elif 'F4' in test and '资金股' in test:
                self._main.TypeKeys("{F4}")
                if left_menus_check():
                    break
            elif 'F5' in test:
                self._main.TypeKeys("{F5}")
                if left_menus_check():
                    break
            elif 'F6' in test:
                self._main.TypeKeys("{F6}")
                if left_menus_check():
                    break
            else:
                try:
                    self._left_treeview.Select(path)                   
                except Exception:
                    pass
                if left_menus_check():
                    break 

    def _bring_main_foreground(self):
        self._main.Minimize()
        time.sleep(0.02)
        self._main.Restore()
        time.sleep(0.02)
        shell = win32com.client.Dispatch("WScript.Shell")
        time.sleep(0.02)
        shell.SendKeys('%')
        time.sleep(0.01)
        pywinauto.win32functions.SetForegroundWindow(self._main.wrapper_object())    

    @property
    def balance(self):
        self._switch_left_menus(self._config.BALANCE_MENU_PATH)
        return self._get_balance_from_statics()

    def _get_balance_from_statics(self):
        result = {}
        for key, control_id in self._config.BALANCE_CONTROL_ID_GROUP.items():
            ww = self._pwindow.window(control_id=control_id, class_name="Static")
            count = 0
            for c in range(20):
                try:
                    test = float(ww.window_text())
                    # 如果股票市值为0, 要多试一下!
                    if (key == "股票市值" and abs(test) < 0.0001 and count < 4):
                        time.sleep(0.05)
                        count += 1
                        continue
                    result[key] = test
                    break
                except Exception:
                    time.sleep(0.05)
        return result
    
    # 注意，各大券商此接口重写，统一输出
    @property
    def position(self):
        self._check_top_window()
        for c in range(2):
            self._switch_left_menus(["查询[F4]", "资金股票"])
            test = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
            if isinstance(test, pd.DataFrame):
                test = test.to_dict("records") if len(test) > 0 else []
                break
            else:
                log.warning("读取position grid失败...")
                test = []
              
        return test

    # 注意，各大券商此接口重写，统一输出
    @property
    def today_entrusts(self):
        self._check_top_window()
        for c in range(2):
            self._switch_left_menus(["查询[F4]", "当日委托"])
            test = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
            if isinstance(test, pd.DataFrame):
                test = test.to_dict("records") if len(test) > 0 else []
                break
            else:
                log.warning("读取today_entrusts grid失败...")
                test = []
              
        return test

    # 注意，各大券商此接口重写，统一输出
    @property
    def today_trades(self):
        self._check_top_window()
        for c in range(2):
            self._switch_left_menus(["查询[F4]", "当日成交"])
            test = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
            if isinstance(test, pd.DataFrame):
                test = test.to_dict("records") if len(test) > 0 else []
                break
            else:
                log.warning("读取today_trades grid失败...")
                test = []
              
        return test

    # 注意，各大券商此接口重写，统一输出   
    @property
    def cancel_entrusts(self):
        self._check_top_window()
        self._refresh()
        for c in range(2):
            self._switch_left_menus(["撤单[F3]"])
            test = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
            if isinstance(test, pd.DataFrame):
                test = test.to_dict("records") if len(test) > 0 else []
                break
            else:
                log.warning("读取cancel_entrusts grid失败...")
                test = []
              
        return test
    
    def cancel_entrust(self, entrust_no):
        """entrust_no: str, 本接口尚未严格测试!"""
        self._check_top_window()
        self._refresh()
        test = self.cancel_entrusts
        for i, entrust in enumerate(test):
            if (
                entrust[self._config.CANCEL_ENTRUST_ENTRUST_FIELD]
                == entrust_no
            ):
                self._cancel_entrust_by_double_click(i)
                return self._handle_pop_dialogs()
        else:
            return {"message": "委托单状态错误不能撤单, 该委托单可能已经成交或者已撤"}

    def trade(self, action, security, price, amount, atype='MARKET', ptype='最优五档成交剩余撤销', **kwargs):
        """
        action   : str, 'BUY' or 'SELL'
        security : str, 股票代码
        price    : str, 交易价格
        amount   : str, 交易数量
        atype    : str, 'MARKET' or 'LIMIT'
        ptype    : str, 委托类型
        """
        a = time.time()
        pass
        
        
        
        
        
        
        
    def buy(self, security, price, amount, **kwargs):
        a = time.time()
        
        self._check_top_window()
        self._switch_left_menus(["买入[F1]"])
        print('000 ..._switch_left_menus', time.time()-a)

        return self.trade(security, price, amount)

    def sell(self, security, price, amount, **kwargs):
        a = time.time()
        self._switch_left_menus(["卖出[F2]"])
        print('000 ..._switch_left_menus', time.time()-a)

        return self.trade(security, price, amount)

    def trade(self, security, price, amount):
        a = time.time()
        self._set_trade_params(security, price, amount)
        print('aaa ..._set_trade_params', time.time()-a)
        
        a = time.time()
        self._submit_trade()
        print('bbb ..._submit_trade', time.time()-a)
        
        a = time.time()
        test = self._handle_pop_dialogs(handler_class=pop_dialog_handler.TradePopDialogHandler)
        print('ccc ..._handle_pop_dialogs', time.time()-a)

        return test

    def _click(self, control_id):
        for c in range(5):
            try:
                test = self._main.window(control_id=control_id, class_name="Button")
                # test.wait("exists visible enabled", 0.05)
                test.click()
                break
            except Exception as e:
                print("_click", e)
                self._check_top_window()
                time.sleep(0.1)

    def _submit_trade(self):
        # 等待股东账号出现!
        for c in range(20):
            try:
                sss = time.time()
                selects = self._main.window(
                    control_id=self._config.TOP_TOOLBAR_CONTROL_ID,
                    class_name="ToolbarWindow32",
                ).window(
                    control_id=self._config.TRADE_ACCOUNT_CONTROL_ID,
                    class_name="ComboBox",
                )
                
                account = selects.window_text()
                if len(account) > 5:
                    # print('----> showup account', account, time.time()-sss)
                    break
            except Exception as e:
                print('等待股东账号出现', e)
            zzz = time.time()
            if (zzz-sss) < 0.05:
                time.sleep(0.05-(zzz-sss))
                print("retry 等待股东账号出现")
        
        for c in range(5):
            try:
                test = self._pwindow.window(control_id=self._config.TRADE_SUBMIT_CONTROL_ID, class_name="Button")
                # test.wait("exists visible enabled", 0.05)
                test.click()
                break
            except Exception as e:
                print("submit_click", e)
                self._check_top_window()
                time.sleep(0.1)

    def _set_trade_params(self, security, price, amount):
        code = security[-6:]
        self._type_keys(self._config.TRADE_SECURITY_CONTROL_ID, code)

        self._type_keys(
            self._config.TRADE_PRICE_CONTROL_ID,
            easyutils.round_price_by_code(price, code),
        )
        
        self._type_keys(self._config.TRADE_AMOUNT_CONTROL_ID, str(int(amount)))
        
        self._wait_trade_showup(self._config.TRADE_SECURITY_NAME_ID, "Static")
    
    
    def market_buy(self, security, amount, ttype=u'最优五档成交剩余撤销', **kwargs):
        """
        市价买入
        :param security: 六位证券代码
        :param amount: 交易数量
        :param ttype: 市价委托类型，默认客户端默认选择，*** 深市删除"即时" ***
                     深市可选 ['1-对手方最优价格','2-本方最优价格','3-即时成交剩余撤销','4-最优五档即时成交剩余撤销','5-全额成交或撤销']
                     沪市可选 ['1-最优五档成交剩余撤销','2-最优五档成交剩余转限价']

        :return: {'entrust_no': '委托单号'}
        """
        a = time.time()
        self._switch_left_menus(["市价委托", "买入"])
        print('000 ..._switch_left_menus', time.time()-a)

        return self.market_trade(security, amount, ttype)

    def market_sell(self, security, amount, ttype=u'最优五档成交剩余撤销', **kwargs):
        """
        市价卖出
        :param security: 六位证券代码
        :param amount: 交易数量
        :param ttype: 市价委托类型，默认客户端默认选择，*** 深市删除"即时" ***
                     深市可选 ['1-对手方最优价格','2-本方最优价格','3-即时成交剩余撤销','4-最优五档即时成交剩余撤销','5-全额成交或撤销']
                     沪市可选 ['1-最优五档成交剩余撤销','2-最优五档成交剩余转限价']

        :return: {'entrust_no': '委托单号'}
        """
        a = time.time()
        self._switch_left_menus(["市价委托", "卖出"])
        print('000 ..._switch_left_menus', time.time()-a)

        return self.market_trade(security, amount, ttype)

    def market_trade(self, security, amount, ttype=None, **kwargs):
        """
        市价交易
        :param security: 六位证券代码
        :param amount: 交易数量
        :param ttype: 市价委托类型，默认客户端默认选择，*** 深市删除"即时" ***
                     深市可选 ['1-对手方最优价格','2-本方最优价格','3-即时成交剩余撤销','4-最优五档即时成交剩余撤销','5-全额成交或撤销']
                     沪市可选 ['1-最优五档成交剩余撤销','2-最优五档成交剩余转限价']

        :return: {'entrust_no': '委托单号'}
        """
        a = time.time()
        self._set_market_trade_params(security, amount)
        print('aaa ..._set_market_trade_params', time.time()-a)
        
        a = time.time()
        self._set_market_trade_type(ttype)
        print('bbb ..._set_market_trade_type', time.time()-a)
        
        a = time.time()
        self._submit_trade()
        print('ccc ..._submit_trade', time.time()-a)
        
        a = time.time()
        test = self._handle_pop_dialogs(handler_class=pop_dialog_handler.TradePopDialogHandler)
        print('ddd ..._handle_pop_dialogs', time.time()-a)

        return test

    def _set_market_trade_type(self, ttype):
        """根据选择的市价交易类型选择对应的下拉选项"""     
        if isinstance(ttype, str): 
            ttype = ttype.replace(u"即时", "")
 
        # 确认市价交易类型选项出现!
        selects = self._wait_trade_showup(self._config.TRADE_MARKET_TYPE_CONTROL_ID, "ComboBox")
                 
        # 选择对应的下拉选项   
        for i, text in enumerate(selects.texts()):
            # skip 0 index, because 0 index is current select index
            text = text.replace(u"即时", "")
            if ttype in text:
                # 如果不是默认选项，则选择下拉
                if i != 0:
                    selects.select(i - 1)
                    
                # 确认市价交易的价格出现!
                self._wait_trade_showup(self._config.TRADE_PRICE_CONTROL_ID, "Edit") 
                break
        else:
            print("不支持对应的市价类型: {}".format(ttype), "将采用默认方式!")
            # 确认市价交易的价格出现
            self._wait_trade_showup(self._config.TRADE_PRICE_CONTROL_ID, "Edit") 

            
    def auto_ipo(self):
        for c in range(2):
            self._switch_left_menus(self._config.AUTO_IPO_MENU_PATH)
            test = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
            if isinstance(test, pd.DataFrame):
                stock_list = test.to_dict("records") if len(test) > 0 else []
                break
            else:
                log.warning("读取auto_ipo grid失败...")
                stock_list = []
              
        if len(stock_list) == 0:
            return {"message": "今日无新股"}
        invalid_list_idx = [
            i for i, v in enumerate(stock_list) if v["申购数量"] <= 0
        ]

        if len(stock_list) == len(invalid_list_idx):
            return {"message": "没有发现可以申购的新股"}

        self._click(self._config.AUTO_IPO_SELECT_ALL_BUTTON_CONTROL_ID)
        self.wait(0.1)

        for row in invalid_list_idx:
            self._click_grid_by_row(row)
        self.wait(0.1)

        self._click(self._config.AUTO_IPO_BUTTON_CONTROL_ID)
        self.wait(0.1)

        return self._handle_pop_dialogs()

    def _click_grid_by_row(self, row):
        x = self._config.COMMON_GRID_LEFT_MARGIN
        y = (
            self._config.COMMON_GRID_FIRST_ROW_HEIGHT
            + self._config.COMMON_GRID_ROW_HEIGHT * row
        )
#         self._check_top_window()
        self._main.window(
            control_id=self._config.COMMON_GRID_CONTROL_ID,
            class_name="CVirtualGridCtrl",
        ).click(coords=(x, y))

    def _run_exe_path(self, exe_path):
        return os.path.join(os.path.dirname(exe_path), "xiadan.exe")

    def wait(self, seconds):
        time.sleep(seconds)

    def exit(self):
        self._app.kill()




    def _set_market_trade_params(self, security, amount):
        code = security[-6:]

        self._type_keys(self._config.TRADE_SECURITY_CONTROL_ID, code)

        self._type_keys(self._config.TRADE_AMOUNT_CONTROL_ID, str(int(amount)))
        
        self._wait_trade_showup(self._config.TRADE_SECURITY_NAME_ID, "Static")

    def _wait_trade_showup(self, control_id, class_name):
        """class_name: "Static", "Edit", "ComboBox" """
        flag = False
        for c in range(10):   # 最大等待10s
            try:
                sss = time.time()
                for i in self._pwindow.Children():
                    condition =  ( 
                        i.control_id() == control_id and 
                        i.class_name() == class_name and 
                        len(i.window_text()) > 1 
                    )
                    if condition and class_name != "ComboBox":
                        flag = True
                        return i     
                    elif condition and class_name == "ComboBox" and '最优五档' in ''.join(i.texts()):
                        flag = True
                        return i  
                if flag:
                    break
                else:
                    print('retry _wait_trade_showup')
                    gaps = time.time() - sss
                    if gaps < 0.05:
                        time.sleep(0.05-gaps)
            except Exception as e:
                print('_wait_trade_showup Exception', e)
                gaps = time.time() - sss
                if gaps < 0.05:
                    time.sleep(0.05-gaps)
                
    def _get_grid_data(self, control_id):
        return self._grid_data_get_strategy.get(control_id)

    def _type_keys(self, control_id, text):
        test = self._pwindow.window(control_id=control_id, class_name="Edit")
        for c in range(50):
            try:
                test.SetEditText(text)
                if test.window_text() == text:
                    break
                else:
                    continue
            except Exception as e:
                print('type:', text, e)
        

    def _cancel_entrust_by_double_click(self, row):
        x = self._config.CANCEL_ENTRUST_GRID_LEFT_MARGIN
        y = (
            self._config.CANCEL_ENTRUST_GRID_FIRST_ROW_HEIGHT
            + self._config.CANCEL_ENTRUST_GRID_ROW_HEIGHT * (row + 1)
        )
        self._main.window(
            control_id=self._config.COMMON_GRID_CONTROL_ID,
            class_name="CVirtualGridCtrl",
        ).double_click(coords=(x, y))

    def _refresh(self):
        self._main.TypeKeys("{F5}")
        
                
    def _handle_pop_dialogs(
        self, handler_class=pop_dialog_handler.PopDialogHandler
    ):
        # 最多等待10秒
        for c in range(50):
            try:
                topw_handle = self._main.PopupWindow() 
                if topw_handle != 0:
                    topw = self._main.window(handle=topw_handle)
                    test = topw.window(control_id=self._config.POP_DIALOD_TITLE_CONTROL_ID)
                    title = test.window_text()
                    if len(title) > 0:
                        handler = handler_class(self._app, topw)
                        result = handler.handle(title)
                        if result:
                            return result
                        else:
                            time.sleep(0.2)
                    else:
                        print('get_pop_dialog_title retry')              
                else:
                    """没弹出，再试几下"""
                    print('没弹出窗口')
                    time.sleep(0.2)
                    pass
            except Exception as e:
                print('pop_dialog', e)
                time.sleep(0.2)
                
        return {"success???": "不应该出现这里"}          

    
    
class BaseLoginClientTrader(ClientTrader):
    @abc.abstractmethod
    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """Login Client Trader"""
        pass

    def prepare(
        self,
        config_path=None,
        user=None,
        password=None,
        exe_path=None,
        comm_password=None,
        **kwargs
    ):
        """
        登陆客户端
        :param config_path: 登陆配置文件，跟参数登陆方式二选一
        :param user: 账号
        :param password: 明文密码
        :param exe_path: 客户端路径类似 r'C:\\htzqzyb2\\xiadan.exe', 默认 r'C:\\htzqzyb2\\xiadan.exe'
        :param comm_password: 通讯密码
        :return:
        """
        if config_path is not None:
            account = helpers.file2dict(config_path)
            user = account["user"]
            password = account["password"]
            comm_password = account.get("comm_password")
            exe_path = account.get("exe_path")
        self.login(
            user,
            password,
            exe_path or self._config.DEFAULT_EXE_PATH,
            comm_password,
            **kwargs
        )
