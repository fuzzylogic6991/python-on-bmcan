# coding: utf-8

"""
Ctypes wrapper module for BUSMUST CAN Interface on win32/win64 systems.

Authors: busmust <busmust@126.com>, BUSMUST Co.,Ltd.
"""

# Import Standard Python Modules
# ==============================
import ctypes
import logging
import sys
import time
from datetime import datetime

try:
    # Try builtin Python 3 Windows API
    from _winapi import WaitForSingleObject, INFINITE

    HAS_EVENTS = True
except ImportError:
    try:
        # Try pywin32 package
        from win32event import WaitForSingleObject, INFINITE

        HAS_EVENTS = True
    except ImportError:
        # Use polling instead
        HAS_EVENTS = False

# Import Modules
# ==============
from can import BusABC, Message, CanError
from can.bus import BusState
from can.util import len2dlc, dlc2len
from .exceptions import SmError
from .bmapi import CAN_OBJ, INIT_CONFIG, BoardInfo
from ctypes import *
import os
import threading

Channel1 = c_uint(0)
Channel2 = c_uint(1)
cwdx = os.getcwd()
STATUS_ERR = 0
STATUS_OK = 1
DevType = c_uint
USBCAN1 = DevType(4)
DevIndex = c_uint(0)
musbcanopen = False

class SmCanBus(BusABC):
    __initialized = False
    ecan = None

    @classmethod
    def __init_class__(cls):
        if not SmCanBus.__initialized:
            SmCanBus.__initialized = True
            SmCanBus.ecan = ECAN()

    def __init__(self, channel,
                 fd=True, receive_own_messages=False, listen_only=False,
                 bitrate=500000, data_bitrate=500000,
                 samplepos=75, data_samplepos=75,
                 tres=False,
                 can_filters=None,
                 **kwargs):

        """
        :param int channel:
            The channel index to create this bus with, which is the index to all available ports when enumerating Busmust devices.
            Can also be a string of the channel's full name. i.e. "BM-CANFD-X1-PRO(1234) CH1"
        :param bool fd:
            If CAN-FD frames should be supported.
        :param bool receive_own_messages:
            If Loopback mode should be supported.
        :param bool listen_only:
            If Listen only mode should be supported, this is the same as setting 'state' property to INACTIVE.
        :param int bitrate:
            Bitrate in bits/s.
        :param int data_bitrate:
            Which bitrate to use for data phase in CAN FD.
            Defaults to arbitration bitrate.
        :param int samplepos:
            Sample position (%).
        :param int data_samplepos:
            Data phase sample pos (%) in CAN FD.
        :param bool tres:
            If 120Ohm CAN terminal register should be enabled.
        """
        SmCanBus.__init_class__()

        # infolist = bmapi.BM_ChannelInfoList)TypeDef(
        # numOfInfo = ctypes.c_int(len(infolist.entries))

        # SmCanBus.ecan.BM_Enumerate(ctypes.byref(infolist), ctypes.byref(numOfInfo))
        # if isinstance(channel, int):
        #     if channel < numOfInfo.value:
        #         self._channelinfo = infolist.entries[channel]
        #     else:
        #         raise BmError(bmapi.BM_ERROR_NODRIVER,
        #                       "Channel %d is not connected or is in use by another app." % channel, "SmCanBus.__init__")
        # elif isinstance(channel, str):
        #     for info in infolist.entries:
        #         if info.name.decode() == channel:
        #             self._channelinfo = info
        #             break
        #     else:
        #         raise BmError(bmapi.BM_ERROR_NODRIVER,
        #                       "Channel %s is not connected or is in use by another app." % channel, "SmCanBus.__init__")

        # self._mode = bmapi.BM_CAN_NORMAL_MODE
        # if not fd:
        #     self._mode = bmapi.BM_CAN_CLASSIC_MODE
        # elif receive_own_messages:
        #     self._mode = bmapi.BM_CAN_EXTERNAL_LOOPBACK_MODE
        # elif listen_only:
        #     self._mode = bmapi.BM_CAN_LISTEN_ONLY_MODE
        #
        # self._tres = bmapi.BM_TRESISTOR_DISABLED
        # if tres:
        #     self._tres = bmapi.BM_TRESISTOR_120
        #
        # self._bitrate = bmapi.BM_BitrateTypeDef()
        # self._bitrate.nbitrate = int(bitrate / 1000)
        # self._bitrate.dbitrate = int(data_bitrate / 1000)
        # self._bitrate.nsamplepos = samplepos
        # self._bitrate.dsamplepos = data_samplepos

        # self._handle = bmapi.BM_ChannelHandle()
        # bmapi.BM_OpenEx(
        #     ctypes.byref(self._handle),
        #     ctypes.byref(self._channelinfo),
        #     self._mode,
        #     self._tres,
        #     ctypes.byref(self._bitrate),
        #     ctypes.cast(ctypes.c_void_p(), ctypes.POINTER(bmapi.BM_RxFilterListTypeDef)), 0
        # )
        # SmCanBus.ecan.OpenDevice()
        # self.channel_info = self._channelinfo.name.decode()
        #


        self.channel = channel
        self.bitrate = bitrate
        self.data_bitrate = data_bitrate
        startTimestamp = ctypes.c_uint32()
        self._time_offset = time.time() - startTimestamp.value * 1e-9
        self.logger = logging.getLogger(__name__)
        self.tmp_signal = 0
        self.tmp_count = 3
        self.average_progress = 0
        self.config_msg = []
        super(SmCanBus, self).__init__(channel=channel, can_filters=can_filters, **kwargs)
        # print(f'init  {channel} {bitrate} {data_bitrate}')
        self.caninit(channel, bitrate, data_bitrate)
        time.sleep(0.05)

    def getTiming(self, mbaud):
        if mbaud == 1152000:
            return 0, 0x14
        if mbaud == 800000:
            return 0, 0x16
        if mbaud == 666000:
            return 0x80, 0xb6
        if mbaud == 500000:
            return 0, 0x1c
        if mbaud == 400000:
            return 0x80, 0xfa
        if mbaud == 250000:
            return 0x01, 0x1c
        if mbaud == 200000:
            return 0x81, 0xfa
        if mbaud == 125000:
            return 0x03, 0x1c
        if mbaud == 100000:
            return 0x04, 0x1c
        if mbaud == 80000:
            return 0x83, 0xff
        if mbaud == 50000:
            return 0x09, 0x1c

    def _apply_filters(self, filters):
        pass

    def _recv_internal(self, timeout):
        channel = ctypes.c_uint32()
        timestamp = ctypes.c_uint32()

        while True:
            try:
                self._is_filtered = False
                length, rec, ret = SmCanBus.ecan.Receivce(USBCAN1, DevIndex, self.channel, 1)
            except SmError as exc:
                raise f"连接can盒失败{exc}"
            else:
                msg = Message(
                    timestamp=timestamp.value * 1e-6 + self._time_offset,
                    arbitration_id=rec[0].ID,
                    is_extended_id=bool(rec[0].ExternFlag),
                    is_remote_frame=bool(rec[0].RemoteFlag),
                    is_error_frame=bool(False),
                    is_fd=False,
                    # error_state_indicator=bool(bmmsg.ctrl.rx.ESI),
                    # bitrate_switch=bool(bmmsg.ctrl.rx.BRS),
                    dlc=rec[0].DataLen,
                    data=rec[0].data,
                    channel=1)
                return msg, self._is_filtered

            if end_time is not None and time.time() > end_time:
                return None, self._is_filtered

            # Wait for receive event to occur
            if timeout is None:
                time_left_ms = INFINITE
            else:
                time_left = end_time - time.time()
                time_left_ms = max(0, int(time_left * 1000))

    def send(self, msg, timeout=None):
        # global musbcanopen
        # if (musbcanopen == False):
        #     print("ERROR请先打开设备.......")
        # else:
        canobj = CAN_OBJ()
        canobj.ID = msg.arbitration_id
        canobj.DataLen = int("8")
        c_ubyte_Array_8 = c_ubyte * 8
        canobj.data = c_ubyte_Array_8(*msg.data)
        canobj.RemoteFlag = 0
        canobj.ExternFlag = 0
        tmp_str = f"can send ID: {hex(canobj.ID)} data: " + hex(canobj.data[0]) + " " + hex(
            canobj.data[1]) + " " + hex(
            canobj.data[2]) + " " + hex(
            canobj.data[3]) + " " + hex(canobj.data[4]) + " " + hex(canobj.data[5]) + " " + hex(
            canobj.data[6]) + " " + hex(canobj.data[7])
        # print(tmp_str)

        result = SmCanBus.ecan.Tramsmit(USBCAN1, DevIndex, self.channel, canobj)
        if result != 1:
            raise  Exception(f"发送失败")

    def shutdown(self):
        SmCanBus.ecan.CloseDevice(DeviceType, DeviceIndex, 1)
        print(f"关闭。。。。。。。。。。。")

    @property
    def state(self):
        return self._state

    @state.setter
    def state(self, new_state):
        mode_changed = False
        self._state = new_state
        if new_state is BusState.ACTIVE:
            if self._mode == bmapi.BM_CAN_OFF_MODE or self._mode == bmapi.BM_CAN_LISTEN_ONLY_MODE:
                self._mode = bmapi.BM_CAN_NORMAL_MODE
                mode_changed = True
            else:
                pass  # Do not change (i.e. loopback)
        elif new_state is BusState.PASSIVE:
            # When this mode is set, the CAN controller does not take part on active events (eg. transmit CAN messages)
            # but stays in a passive mode (CAN monitor), in which it can analyse the traffic on the CAN bus used by a BMCAN channel.
            if self._mode != bmapi.BM_CAN_LISTEN_ONLY_MODE:
                self._mode = bmapi.BM_CAN_LISTEN_ONLY_MODE
                mode_changed = True
        if mode_changed:
            bmapi.BM_SetCanMode(self._handle, self._mode)

    @classmethod
    def enumerate(cls):
        SmCanBus.__init_class__()
        infolist = bmapi.BM_ChannelInfoListTypeDef()
        numOfInfo = ctypes.c_int(len(infolist.entries))
        bmapi.BM_Enumerate(ctypes.byref(infolist), ctypes.byref(numOfInfo))
        channellist = []
        for i in range(numOfInfo.value):
            channellist.append({
                'index': i,
                'name': infolist.entries[i].name.decode(),
                # Add other exports here
            })
        return channellist

    def send_isotp(self, payload, timeout=-1):
        timeout_ms = int(timeout * 1000.0) if timeout >= 0 else -1
        bmapi.BM_WriteIsotp(self._handle, ctypes.c_char_p(payload), len(payload), timeout_ms,
                            ctypes.byref(self._isotp_config))

    def caninit(self, chan, bitrate, data_bitrate):
        initconfig = INIT_CONFIG()
        initconfig.acccode = 0  # 设置验收码
        initconfig.accmask = 0xFFFFFFFF  # 设置屏蔽码
        initconfig.filter = 0  # 设置滤波使能
        initconfig.timing0, initconfig.timing1 = self.getTiming(bitrate)
        # print(f'initconfig {initconfig.timing0} {initconfig.timing1}')
        initconfig.mode = 0
        if (SmCanBus.ecan.InitCan(USBCAN1, DevIndex, chan, initconfig) != STATUS_OK):
            msg = f"ERROR InitCan {chan} Failed!"
            self.logger.error(msg)
            raise Exception(msg)
        if (SmCanBus.ecan.StartCan(USBCAN1, DevIndex, chan) != STATUS_OK):
            msg = f"ERROR StartCan {chan} Failed!"
            self.logger.error(msg)
            raise Exception(msg)

    def readmess(self):
        # global musbcanopen
        # if (musbcanopen == False):
        #     self.logger.error("ERROR 请先打开设备")
        # else:
        mboardinfo, ret = SmCanBus.ecan.ReadBoardInfo(USBCAN1, c_uint(0))  # 读取设备信息需要在打开设备后执行
        if ret == STATUS_OK:
            mstr = ""
            for i in range(0, 10):
                mstr = mstr + chr(mboardinfo.str_Serial_Num[i])  # 结构体中str_Serial_Num内部存放存放SN号的ASC码
            self.logger.info("SN:" + mstr)


class ECAN(object):
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        try:
            current_dir = os.getcwd()
            dll_path = os.path.join(current_dir, 'ECanVci64.dll')
            self.dll = cdll.LoadLibrary(dll_path)
        except Exception as e:
            print("DLL Couldn't be loaded:%s", e)
        if self.dll == None:
            print("DLL Couldn't be loaded")

        if (self.OpenDevice(USBCAN1, c_uint(0)) != STATUS_OK):
            msg = f"ERROR OpenDevice Failed!"
            self.logger.error(msg)
            raise Exception(msg)

    def OpenDevice(self, DeviceType, DeviceIndex):
        try:
            return self.dll.OpenDevice(DeviceType, DeviceIndex, 0)
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on OpenDevice!")
            raise

    def CloseDevice(self, DeviceType, DeviceIndex):
        try:
            return self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
        except:
            self.logger.error("Exception on CloseDevice!")
            raise

    def ReadBoardInfo(self, DeviceType, DeviceIndex):
        try:
            mboardinfo = BoardInfo()
            ret = self.dll.ReadBoardInfo(DeviceType, DeviceIndex, byref(mboardinfo))
            return mboardinfo, ret
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on ReadBoardInfo!")
            raise

    def Receivce(self, DeviceType, DeviceIndex, CanInd, lenthgth):
        try:
            recmess = (CAN_OBJ * lenthgth)()
            ret = self.dll.Receive(DeviceType, DeviceIndex, CanInd, byref(recmess), lenthgth, 0)
            return lenthgth, recmess, ret
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on Receive!")
            raise

    def Tramsmit(self, DeviceType, DeviceIndex, CanInd, mcanobj):
        try:
            return self.dll.Transmit(DeviceType, DeviceIndex, CanInd, byref(mcanobj), c_uint16(1))
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on Tramsmit!")
            raise

    def StartCan(self, DeviceType, DeviceIndex, CanInd):
        try:
            return self.dll.StartCAN(DeviceType, DeviceIndex, CanInd)
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on StartCan!")
            raise

    def InitCan(self, DeviceType, DeviceIndex, CanInd, Initconfig):
        try:
            return self.dll.InitCAN(DeviceType, DeviceIndex, CanInd, byref(Initconfig))
        except:
            self.dll.CloseDevice(DeviceType, DeviceIndex, 0)
            self.logger.error("Exception on InitCan!")
            raise
