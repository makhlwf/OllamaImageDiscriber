# -*- coding: utf-8 -*-
import sys
import os
import io
import json
import threading
import base64
import urllib.request
import urllib.error
import socket
import tempfile
import wx

import api
import ui
import core
import config
import gui
import globalPluginHandler
import scriptHandler
import queueHandler
import addonHandler
import tones
from logHandler import log

addonHandler.initTranslation()

CONF_SECTION = "ollamaDescriber"
DEFAULT_MODEL = "qwen3.5:0.8b"
DEFAULT_PROMPT = "Describe this UI element or image concisely."
DEFAULT_HOST = "http://localhost:11434"

# Ensure config exists
conf = config.conf
if CONF_SECTION not in conf:
    conf[CONF_SECTION] = {}

# Smart Config Value Parser
def get_config(key, default):
    return conf[CONF_SECTION].get(key, default)

def get_config_int(key, default):
    try:
        return int(conf[CONF_SECTION].get(key, default))
    except (ValueError, TypeError):
        return default

def get_config_bool(key, default):
    val = conf[CONF_SECTION].get(key, default)
    if isinstance(val, str):
        return val.lower() == "true"
    return bool(val)


class OllamaSettingsPanel(gui.settingsDialogs.SettingsPanel):
    title = _("Ollama Image Describer")

    def makeSettings(self, settingsSizer):
        sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

        self.modelControl = sHelper.addLabeledControl(_("Model Name:"), wx.TextCtrl)
        self.modelControl.SetValue(get_config("model", DEFAULT_MODEL))

        self.promptControl = sHelper.addLabeledControl(_("Default System Prompt:"), wx.TextCtrl)
        self.promptControl.SetValue(get_config("prompt", DEFAULT_PROMPT))
        
        self.hostControl = sHelper.addLabeledControl(_("Ollama Server URL:"), wx.TextCtrl)
        self.hostControl.SetValue(get_config("host", DEFAULT_HOST))

        self.apiControl = sHelper.addLabeledControl(_("API Key (Optional):"), wx.TextCtrl)
        self.apiControl.SetValue(get_config("apikey", ""))
        
        self.timeoutControl = sHelper.addLabeledControl(_("Timeout (seconds):"), wx.SpinCtrl, min=5, max=120)
        self.timeoutControl.SetValue(get_config_int("timeout", 30))

        self.maxSizeControl = sHelper.addLabeledControl(_("Max Image Size (px):"), wx.SpinCtrl, min=256, max=4096)
        self.maxSizeControl.SetValue(get_config_int("maxImageSize", 1024))

        self.virtualViewerCheckbox = sHelper.addItem(wx.CheckBox(self, label=_("Show description in Virtual Viewer")))
        self.virtualViewerCheckbox.SetValue(get_config_bool("useVirtualViewer", True))

        self.clipboardCheckbox = sHelper.addItem(wx.CheckBox(self, label=_("Automatically copy description to clipboard")))
        self.clipboardCheckbox.SetValue(get_config_bool("copyToClipboard", False))

        self.debugCheckbox = sHelper.addItem(wx.CheckBox(self, label=_("Save captured image to Temp folder (Debug)")))
        self.debugCheckbox.SetValue(get_config_bool("debugSave", False))

    def onSave(self):
        conf[CONF_SECTION]["model"] = self.modelControl.GetValue()
        conf[CONF_SECTION]["prompt"] = self.promptControl.GetValue()
        conf[CONF_SECTION]["host"] = self.hostControl.GetValue().rstrip('/')
        conf[CONF_SECTION]["apikey"] = self.apiControl.GetValue()
        conf[CONF_SECTION]["timeout"] = self.timeoutControl.GetValue()
        conf[CONF_SECTION]["maxImageSize"] = self.maxSizeControl.GetValue()
        conf[CONF_SECTION]["useVirtualViewer"] = self.virtualViewerCheckbox.GetValue()
        conf[CONF_SECTION]["copyToClipboard"] = self.clipboardCheckbox.GetValue()
        conf[CONF_SECTION]["debugSave"] = self.debugCheckbox.GetValue()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    
    def __init__(self):
        super(GlobalPlugin, self).__init__()
        self.is_processing = False
        self.cancel_event = threading.Event()
        self.last_response = ""
        
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(OllamaSettingsPanel)

    def terminate(self):
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(OllamaSettingsPanel)
        super(GlobalPlugin, self).terminate()

    def play_sound(self, sound_type):
        if sound_type == "start":
            tones.beep(800, 50)
        elif sound_type == "success":
            tones.beep(1200, 100)
        elif sound_type == "error":
            tones.beep(300, 200)

    def _processing_heartbeat(self):
        if self.is_processing and not self.cancel_event.is_set():
            tones.beep(500, 20)
            core.callLater(2000, self._processing_heartbeat)

    def process_wx_image(self, image):
        width = image.GetWidth()
        height = image.GetHeight()
        max_dim = get_config_int("maxImageSize", 1024)

        if width > max_dim or height > max_dim:
            ratio = min(max_dim / width, max_dim / height)
            new_w, new_h = max(1, int(width * ratio)), max(1, int(height * ratio))
            image.Rescale(new_w, new_h, wx.IMAGE_QUALITY_HIGH)

        stream = io.BytesIO()
        image.SaveFile(stream, wx.BITMAP_TYPE_JPEG)
        img_data = stream.getvalue()
        stream.close()

        if get_config_bool("debugSave", False):
            temp_path = os.path.join(tempfile.gettempdir(), "nvda_ollama_debug.jpg")
            with open(temp_path, "wb") as f:
                f.write(img_data)
            log.info(f"Saved debug image to {temp_path}")

        return img_data

    def get_clipboard_image(self):
        img_data = None
        if wx.TheClipboard.Open():
            try:
                if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_BITMAP)):
                    data = wx.BitmapDataObject()
                    wx.TheClipboard.GetData(data)
                    bmp = data.GetBitmap()
                    if bmp.IsOk():
                        image = bmp.ConvertToImage()
                        img_data = self.process_wx_image(image)
            finally:
                wx.TheClipboard.Close()
        return img_data

    def take_screenshot(self, obj=None, full_screen=False):
        try:
            if full_screen:
                x, y = 0, 0
                width, height = api.getDesktopObject().location[2], api.getDesktopObject().location[3]
            else:
                location = obj.location
                if not location:
                    return None
                x, y, width, height = location
            
            if width <= 0 or height <= 0:
                return None

            bmp = wx.Bitmap(width, height)
            mem_dc = wx.MemoryDC(bmp)
            screen_dc = wx.ScreenDC()
            mem_dc.Blit(0, 0, width, height, screen_dc, x, y)
            mem_dc.SelectObject(wx.NullBitmap)
            
            image = bmp.ConvertToImage()
            return self.process_wx_image(image)
        except Exception as e:
            log.error(f"Ollama Screenshot Error: {e}")
            return None

    def worker_process_image(self, img_data, custom_prompt=None, context_info=""):
        try:
            model = get_config("model", DEFAULT_MODEL)
            base_prompt = custom_prompt if custom_prompt else get_config("prompt", DEFAULT_PROMPT)
            
            if context_info:
                prompt = f"[System UI Context: {context_info}]\n\n{base_prompt}"
            else:
                prompt = base_prompt

            host = get_config("host", DEFAULT_HOST).rstrip('/')
            api_key = get_config("apikey", "")
            timeout = get_config_int("timeout", 30)
            
            b64_image = base64.b64encode(img_data).decode('utf-8')

            payload = {
                "model": model,
                "stream": False,
                "messages":[
                    {
                        "role": "user",
                        "content": prompt,
                        "images":[b64_image]
                    }
                ]
            }

            url = f"{host}/api/chat"
            data = json.dumps(payload).encode('utf-8')
            
            headers = {'Content-Type': 'application/json'}
            if api_key:
                headers['Authorization'] = f"Bearer {api_key}"

            req = urllib.request.Request(url, data=data, headers=headers)
            
            with urllib.request.urlopen(req, timeout=timeout) as response:
                if self.cancel_event.is_set():
                    return

                if response.getcode() != 200:
                    raise urllib.error.URLError(f"HTTP Error {response.getcode()}")

                raw_data = response.read().decode('utf-8')
                result = json.loads(raw_data)
                content = result.get('message', {}).get('content', '')
                
                if content:
                    self.play_sound("success")
                    core.callLater(10, self.handle_success, content)
                else:
                    self.play_sound("error")
                    core.callLater(10, ui.message, _("Ollama returned an empty response."))

        except json.JSONDecodeError:
            self.play_sound("error")
            core.callLater(10, ui.message, _("Invalid response from server. Check Ollama host url."))
            log.error("Ollama API Error: JSON Decode Failed")
        except socket.timeout:
            if not self.cancel_event.is_set():
                self.play_sound("error")
                core.callLater(10, ui.message, _("Ollama server timed out. Try increasing the timeout."))
        except urllib.error.URLError as e:
            if not self.cancel_event.is_set():
                self.play_sound("error")
                core.callLater(10, ui.message, _("Connection failed. Is Ollama running?"))
                log.error(f"Ollama Connection Error: {e}")
        except Exception as e:
            if not self.cancel_event.is_set():
                self.play_sound("error")
                core.callLater(10, ui.message, _("An unexpected Ollama Error occurred."))
                log.error(f"Ollama API Error: {e}")
        finally:
            self.is_processing = False

    def handle_success(self, content):
        self.last_response = content
        
        if get_config_bool("copyToClipboard", False):
            api.copyToClip(content)
            ui.message(_("Copied to clipboard."))

        if get_config_bool("useVirtualViewer", True):
            ui.browseableMessage(content, _("Ollama Description"))
        else:
            ui.message(content)

    def trigger_capture(self, obj=None, full_screen=False, from_clipboard=False, custom_prompt=None):
        if self.is_processing:
            ui.message(_("Already analyzing an image. Please wait or cancel."))
            return

        ui.message(_("Capturing..."))
        self.play_sound("start")
        
        context_info = ""
        if from_clipboard:
            img_data = self.get_clipboard_image()
            if not img_data:
                self.play_sound("error")
                ui.message(_("No valid image found on the clipboard."))
                return
        else:
            img_data = self.take_screenshot(obj, full_screen)
            if obj:
                role = obj.roleText if hasattr(obj, 'roleText') else "Element"
                name = obj.name if hasattr(obj, 'name') and obj.name else "Unknown"
                context_info = f"Type: {role}, Name: '{name}'"

        if img_data:
            self.is_processing = True
            self.cancel_event.clear()
            core.callLater(100, self._processing_heartbeat)
            
            t = threading.Thread(target=self.worker_process_image, args=(img_data, custom_prompt, context_info))
            t.daemon = True
            t.start()
        else:
            self.play_sound("error")
            ui.message(_("Could not capture image. Object might be invisible or off-screen."))

    def _prompt_and_capture(self, obj=None, full_screen=False, from_clipboard=False):
        def show_dialog():
            dlg = wx.TextEntryDialog(gui.mainFrame, _("What would you like to ask about this image?"), _("Ask Ollama"))
            gui.mainFrame.prePopup()
            try:
                if dlg.ShowModal() == wx.ID_OK:
                    prompt = dlg.GetValue()
                    if prompt.strip():
                        core.callLater(10, self.trigger_capture, obj, full_screen, from_clipboard, prompt)
            finally:
                gui.mainFrame.postPopup()
                dlg.Destroy()
        wx.CallAfter(show_dialog)

    # SCRIPT HANDLERS
    
    def script_describeObject(self, gesture):
        self.trigger_capture(obj=api.getNavigatorObject())
    script_describeObject.__doc__ = _("Sends an image of the current navigator object to Ollama.")
    script_describeObject.category = _("Ollama Image Describer")
    
    def script_describeScreen(self, gesture):
        self.trigger_capture(full_screen=True)
    script_describeScreen.__doc__ = _("Sends an image of the entire screen to Ollama.")
    script_describeScreen.category = _("Ollama Image Describer")

    def script_describeClipboard(self, gesture):
        self.trigger_capture(from_clipboard=True)
    script_describeClipboard.__doc__ = _("Sends an image from the Windows clipboard to Ollama.")
    script_describeClipboard.category = _("Ollama Image Describer")

    def script_askObject(self, gesture):
        if self.is_processing:
            ui.message(_("Already analyzing. Please wait."))
            return
        self._prompt_and_capture(obj=api.getNavigatorObject())
    script_askObject.__doc__ = _("Prompts for a custom question about the current object.")
    script_askObject.category = _("Ollama Image Describer")

    def script_askScreen(self, gesture):
        if self.is_processing:
            ui.message(_("Already analyzing. Please wait."))
            return
        self._prompt_and_capture(full_screen=True)
    script_askScreen.__doc__ = _("Prompts for a custom question about the whole screen.")
    script_askScreen.category = _("Ollama Image Describer")

    def script_cancelRequest(self, gesture):
        if self.is_processing:
            self.cancel_event.set()
            self.is_processing = False
            self.play_sound("error")
            ui.message(_("Ollama request cancelled."))
        else:
            ui.message(_("No Ollama request running."))
    script_cancelRequest.__doc__ = _("Cancels the ongoing Ollama analysis request.")
    script_cancelRequest.category = _("Ollama Image Describer")

    def script_repeatLastResponse(self, gesture):
        if self.last_response:
            if get_config_bool("useVirtualViewer", True):
                ui.browseableMessage(self.last_response, _("Last Ollama Description"))
            else:
                ui.message(self.last_response)
        else:
            ui.message(_("No previous description found."))
    script_repeatLastResponse.__doc__ = _("Repeats the last Ollama description text.")
    script_repeatLastResponse.category = _("Ollama Image Describer")

    __gestures = {
        "kb:NVDA+Control+Shift+O": "describeObject",
        "kb:NVDA+Windows+Shift+O": "describeScreen",
        "kb:NVDA+Control+Windows+O": "describeClipboard",
        "kb:NVDA+Alt+Shift+O": "askObject",
        "kb:NVDA+Alt+Windows+O": "askScreen",
        "kb:NVDA+Shift+Windows+C": "cancelRequest",
        "kb:NVDA+Shift+Windows+R": "repeatLastResponse"
    }