"""
OBS Audio Capture Alert Plugin (Modified)
Monitors audio input devices and sends Windows toast notification when silence is detected.

Original: https://obsproject.com/forum/resources/audio-capture-alert.2063/
Modified: Added Windows toast notification support

Features:
- Configurable silence threshold (in dB and duration)
- Windows toast notification when silence detected
- Optional: Enable image, media or video sources
- Real-time audio level monitoring
- Configurable check interval timer
- Notification cooldown to prevent spam
"""

import obspython as obs
from types import SimpleNamespace
from ctypes import *
from ctypes.util import find_library
import math
import subprocess
import threading

# Load the OBS library
obsffi = CDLL(find_library("obs"))
G = SimpleNamespace()

def wrap(funcname, restype, argtypes):
    """Simplify wrapping ctypes functions in obsffi"""
    func = getattr(obsffi, funcname)
    func.restype = restype
    func.argtypes = argtypes
    globals()["g_" + funcname] = func

class Source(Structure):
    pass

class Volmeter(Structure):
    pass

# Define the callback type for the volmeter
volmeter_callback_t = CFUNCTYPE(None, c_void_p, POINTER(c_float), POINTER(c_float), POINTER(c_float))

# Wrap OBS functions
wrap("obs_get_source_by_name", POINTER(Source), argtypes=[c_char_p])
wrap("obs_source_release", None, argtypes=[POINTER(Source)])
wrap("obs_volmeter_create", POINTER(Volmeter), argtypes=[c_int])
wrap("obs_volmeter_destroy", None, argtypes=[POINTER(Volmeter)])
wrap("obs_volmeter_add_callback", None, argtypes=[POINTER(Volmeter), volmeter_callback_t, c_void_p])
wrap("obs_volmeter_remove_callback", None, argtypes=[POINTER(Volmeter), volmeter_callback_t, c_void_p])
wrap("obs_volmeter_attach_source", c_bool, argtypes=[POINTER(Volmeter), POINTER(Source)])

# Volmeter callback function
@volmeter_callback_t
def volmeter_callback(data, mag, peak, input):
    G.noise = float(peak[0])  # Peak volume in dB

# Constants and global variables
OBS_FADER_LOG = 2
G.lock = False
G.start_delay = 1  # Delay before starting to monitor
G.duration = 0
G.noise = -math.inf  # Default value for noise (silence)
G.tick = 10000  # Default timer tick in milliseconds (10 seconds)
G.tick_mili = G.tick * 0.001
G.mic_source_name = ""  # Name of the audio capture source
G.image_source_name = ""  # Name of the image source to enable
G.media_source_name = ""  # Name of the media source to enable
G.video_source_name = ""  # Name of the video capture device to enable
G.volmeter = None  # Placeholder for the volmeter instance
G.silence_duration = 0  # Duration of silence in seconds
G.silence_threshold = 60  # Default silence threshold in seconds (1 minute)
G.silence_db_threshold = -60  # Silence threshold in dB (adjust as needed)
G.plugin_enabled = False  # Plugin disabled by default
G.enable_only_active = False  # Only enable when streaming/recording
G.event_logging = False  # Event logging disabled by default

# New settings for Windows notification
G.enable_windows_notification = True  # Enable Windows toast notification
G.notification_title = "OBS マイク警告"
G.notification_message = "マイクの音が入っていないかも？"
G.notification_cooldown = 60  # Cooldown between notifications (seconds)
G.last_notification_time = 0  # Track last notification time
G.notification_sent = False  # Track if notification was sent for current silence period
G.enable_obs_source = False  # Option to also enable OBS source
G.prev_output_active = False  # Track previous streaming/recording state

def send_windows_notification(title, message):
    """Send Windows toast notification using PowerShell (no additional packages required)"""
    def _send():
        try:
            # PowerShell script for toast notification
            ps_script = f'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime] | Out-Null

$template = @"
<toast duration="short">
    <visual>
        <binding template="ToastText02">
            <text id="1">{title}</text>
            <text id="2">{message}</text>
        </binding>
    </visual>
    <audio src="ms-winsoundevent:Notification.Default"/>
</toast>
"@

$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($template)
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("OBS Studio").Show($toast)
'''
            # Run PowerShell in hidden window
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            
            subprocess.run(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_script],
                startupinfo=startupinfo,
                capture_output=True,
                timeout=10
            )
            if G.event_logging:
                print(f"Windows notification sent: {title} - {message}")
        except Exception as e:
            print(f"Failed to send Windows notification: {e}")
    
    # Run in separate thread to avoid blocking OBS
    thread = threading.Thread(target=_send, daemon=True)
    thread.start()

class _Functions:
    def __init__(self, source_name=None):
        self.source_name = source_name

    def set_visible_all(self, visible):
        """Cycle through all scenes, manually toggling visibility of the source"""
        if G.event_logging:
            print(f"Attempting to set visibility of '{self.source_name}' to {visible}")
        scenes = obs.obs_frontend_get_scenes()
        if not scenes:
            if G.event_logging:
                print("No scenes found!")
            return
        for scene in scenes:
            scene_test = obs.obs_scene_from_source(scene)
            if not scene_test:
                if G.event_logging:
                    print(f"Failed to get scene from source")
                continue
            in_scene = obs.obs_scene_find_source(scene_test, self.source_name)
            if in_scene:
                if G.event_logging:
                    print(f"Found source '{self.source_name}' in scene '{obs.obs_source_get_name(scene)}'")
                obs.obs_sceneitem_set_visible(in_scene, visible)
                if G.event_logging:
                    print(f"Set visibility of '{self.source_name}' to {visible}")
            else:
                if G.event_logging:
                    print(f"Source '{self.source_name}' not found in scene '{obs.obs_source_get_name(scene)}'")
        obs.source_list_release(scenes)

# Event loop for monitoring audio levels
def event_loop():
    """Check audio levels every tick interval."""
    global G
    
    # Check if plugin should be active based on streaming/recording state
    if G.enable_only_active:
        output_active = obs.obs_frontend_streaming_active() or obs.obs_frontend_recording_active()

        # Reset state when output state changes (start or stop)
        if G.prev_output_active != output_active:
            G.silence_duration = 0
            G.notification_sent = False
            if G.event_logging:
                if output_active:
                    print("Recording/streaming started - reset notification state")
                else:
                    print("Recording/streaming stopped - reset notification state")

        G.prev_output_active = output_active

        if not output_active:
            if G.event_logging:
                print("Not streaming or recording - plugin inactive")
            return
    
    if G.event_logging:
        print(f"G.noise = {G.noise} dB (Silence Duration: {G.silence_duration}s)")
    if G.duration > G.start_delay:
        if not G.lock:
            if G.event_logging:
                print("Initializing volmeter...")
            source = g_obs_get_source_by_name(G.mic_source_name.encode("utf-8"))
            if not source:
                print(f"Error: Audio Capture source '{G.mic_source_name}' not found!")
                return
            G.volmeter = g_obs_volmeter_create(OBS_FADER_LOG)
            if not G.volmeter:
                print("Error: Failed to create volmeter!")
                g_obs_source_release(source)
                return
            g_obs_volmeter_add_callback(G.volmeter, volmeter_callback, None)
            if g_obs_volmeter_attach_source(G.volmeter, source):
                g_obs_source_release(source)
                G.lock = True
                if G.event_logging:
                    print("Volmeter attached to Audio Capture source.")
            else:
                print("Error: Failed to attach volmeter to Audio Capture source!")
                g_obs_volmeter_destroy(G.volmeter)
                g_obs_source_release(source)
                return
        # Check for silence
        if G.noise <= G.silence_db_threshold or math.isinf(G.noise):  # Silence or -inf
            G.silence_duration += G.tick / 1000  # Increment silence duration by tick interval in seconds
            if G.silence_duration >= G.silence_threshold:
                if G.event_logging:
                    print(f"Silence detected for {G.silence_threshold} seconds.")
                
                # Send Windows notification (only once per silence period)
                if G.enable_windows_notification and not G.notification_sent:
                    send_windows_notification(G.notification_title, G.notification_message)
                    G.notification_sent = True
                
                # Enable OBS source if option is enabled
                if G.enable_obs_source:
                    enable_source(True)
        else:
            # Sound detected - reset
            if G.silence_duration > 0:
                if G.event_logging:
                    print("Sound detected - resetting silence counter")
            G.silence_duration = 0
            G.notification_sent = False  # Reset notification flag
            if G.enable_obs_source:
                enable_source(False)
    else:
        G.duration += G.tick_mili  # Increment duration by tick interval

def enable_source(enable):
    """Enable or disable the specified source (image, media or video)."""
    source_name = G.image_source_name or G.media_source_name or G.video_source_name
    if not source_name:
        if G.event_logging:
            print("No OBS source selected for alert")
        return
    # Use the _Functions class to set visibility
    func = _Functions(source_name)
    func.set_visible_all(enable)
    if G.event_logging:
        print(f"Source '{source_name}' {'enabled' if enable else 'disabled'}")

def on_frontend_event(event):
    """Handle OBS frontend events for recording/streaming start/stop."""
    if not G.enable_only_active:
        return

    # Reset state when recording or streaming starts/stops
    if event in (obs.OBS_FRONTEND_EVENT_RECORDING_STARTED,
                 obs.OBS_FRONTEND_EVENT_RECORDING_STOPPED,
                 obs.OBS_FRONTEND_EVENT_STREAMING_STARTED,
                 obs.OBS_FRONTEND_EVENT_STREAMING_STOPPED):
        G.silence_duration = 0
        G.notification_sent = False
        if G.event_logging:
            event_names = {
                obs.OBS_FRONTEND_EVENT_RECORDING_STARTED: "Recording started",
                obs.OBS_FRONTEND_EVENT_RECORDING_STOPPED: "Recording stopped",
                obs.OBS_FRONTEND_EVENT_STREAMING_STARTED: "Streaming started",
                obs.OBS_FRONTEND_EVENT_STREAMING_STOPPED: "Streaming stopped",
            }
            print(f"{event_names.get(event, 'Unknown event')} - reset notification state")

def script_description():
    return """<h2>Audio Capture Alert (Modified)</h2>
<p>マイクの音声を監視し、一定時間無音が続いたらWindows通知でお知らせします。</p>
<p><b>使い方:</b></p>
<ol>
<li>監視するオーディオソースを選択</li>
<li>無音判定の時間を設定</li>
<li>通知メッセージをカスタマイズ（任意）</li>
<li>プラグインを有効化</li>
</ol>
"""

def script_load(settings):
    """Called when the script is loaded."""
    obs.obs_frontend_add_event_callback(on_frontend_event)

def script_unload():
    # Remove frontend event callback
    obs.obs_frontend_remove_event_callback(on_frontend_event)
    # Remove timer
    obs.timer_remove(event_loop)
    # Clean up volmeter
    if G.volmeter:
        g_obs_volmeter_remove_callback(G.volmeter, volmeter_callback, None)
        g_obs_volmeter_destroy(G.volmeter)
        print("Volmeter and callback removed.")
    else:
        print("No volmeter to clean up.")

def script_defaults(settings):
    obs.obs_data_set_default_int(settings, "tick_interval", 10)  # Default to 10 seconds
    obs.obs_data_set_default_int(settings, "silence_threshold", 30)  # Default to 30 seconds
    obs.obs_data_set_default_bool(settings, "enable_windows_notification", True)
    obs.obs_data_set_default_string(settings, "notification_title", "OBS マイク警告")
    obs.obs_data_set_default_string(settings, "notification_message", "マイクの音が入っていないかも？")
    obs.obs_data_set_default_bool(settings, "enable_obs_source", False)
    obs.obs_data_set_default_bool(settings, "enable_only_active", True)  # Default to ON

def script_properties():
    props = obs.obs_properties_create()
    
    # ===== Audio Capture Source Section =====
    obs.obs_properties_add_text(props, "section1", "【監視するオーディオソース】", obs.OBS_TEXT_INFO)
    mic_list = obs.obs_properties_add_list(props, "mic_source_name", "オーディオソース",
        obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(mic_list, "-- 選択してください --", "")
    
    # ===== Timer Settings Section =====
    obs.obs_properties_add_text(props, "section2", "【タイマー設定】", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_int(props, "tick_interval", "チェック間隔（秒）", 1, 60, 1)
    obs.obs_properties_add_int(props, "silence_threshold", "無音判定時間（秒）", 5, 600, 5)
    
    # ===== Windows Notification Section =====
    obs.obs_properties_add_text(props, "section3", "【Windows通知設定】", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_bool(props, "enable_windows_notification", "Windows通知を有効にする")
    obs.obs_properties_add_text(props, "notification_title", "通知タイトル", obs.OBS_TEXT_DEFAULT)
    obs.obs_properties_add_text(props, "notification_message", "通知メッセージ", obs.OBS_TEXT_DEFAULT)
    
    # ===== OBS Source Alert Section (Optional) =====
    obs.obs_properties_add_text(props, "section4", "【OBSソース表示（オプション）】", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_bool(props, "enable_obs_source", "OBSソースも表示する")
    source_list = obs.obs_properties_add_list(props, "combined_source", "アラートソース",
        obs.OBS_COMBO_TYPE_LIST, obs.OBS_COMBO_FORMAT_STRING)
    obs.obs_property_list_add_string(source_list, "-- 選択してください --", "")
    
    # ===== Plugin Control Section =====
    obs.obs_properties_add_text(props, "section5", "【プラグイン制御】", obs.OBS_TEXT_INFO)
    obs.obs_properties_add_bool(props, "plugin_enabled", "プラグインを有効にする")
    obs.obs_properties_add_bool(props, "enable_only_active", "配信/録画中のみ有効")
    obs.obs_properties_add_bool(props, "event_logging", "デバッグログを有効にする")

    # Populate dropdowns with available sources
    sources = obs.obs_enum_sources()
    if sources:
        for source in sources:
            source_id = obs.obs_source_get_id(source)
            name = obs.obs_source_get_name(source)
            # Add audio sources to the audio capture dropdown
            if source_id in ["wasapi_input_capture", "wasapi_output_capture", "coreaudio_input_capture", "dshow_input", "pulse_input_capture", "alsa_input_capture"]:
                obs.obs_property_list_add_string(mic_list, name, name)
            # Add image, media and video sources to combined dropdown
            if source_id == "image_source":
                obs.obs_property_list_add_string(source_list, f"[画像] {name}", f"image:{name}")
            elif source_id == "ffmpeg_source":
                obs.obs_property_list_add_string(source_list, f"[メディア] {name}", f"media:{name}")
            elif source_id == "dshow_input":
                obs.obs_property_list_add_string(source_list, f"[映像] {name}", f"video:{name}")
        obs.source_list_release(sources)
    return props

def script_update(settings):
    G.mic_source_name = obs.obs_data_get_string(settings, "mic_source_name")
    
    # Parse combined source selection
    combined_source = obs.obs_data_get_string(settings, "combined_source")
    if combined_source and combined_source.startswith("image:"):
        G.image_source_name = combined_source[6:]
        G.media_source_name = ""
        G.video_source_name = ""
    elif combined_source and combined_source.startswith("media:"):
        G.media_source_name = combined_source[6:]
        G.image_source_name = ""
        G.video_source_name = ""
    elif combined_source and combined_source.startswith("video:"):
        G.video_source_name = combined_source[6:]
        G.image_source_name = ""
        G.media_source_name = ""
    else:
        G.image_source_name = ""
        G.media_source_name = ""
        G.video_source_name = ""
    
    G.tick = (obs.obs_data_get_int(settings, "tick_interval") or 10) * 1000
    G.silence_threshold = obs.obs_data_get_int(settings, "silence_threshold") or 30
    
    # Windows notification settings
    G.enable_windows_notification = obs.obs_data_get_bool(settings, "enable_windows_notification")
    G.notification_title = obs.obs_data_get_string(settings, "notification_title") or "OBS マイク警告"
    G.notification_message = obs.obs_data_get_string(settings, "notification_message") or "マイクの音が入っていないかも？"
    G.enable_obs_source = obs.obs_data_get_bool(settings, "enable_obs_source")
    
    # Get current values before updating
    prev_plugin_enabled = G.plugin_enabled
    prev_enable_only_active = G.enable_only_active
    
    # Update settings
    G.plugin_enabled = obs.obs_data_get_bool(settings, "plugin_enabled")
    G.enable_only_active = obs.obs_data_get_bool(settings, "enable_only_active")
    G.event_logging = obs.obs_data_get_bool(settings, "event_logging")
    
    # Reset silence duration if plugin enabled state changed
    if prev_plugin_enabled != G.plugin_enabled or prev_enable_only_active != G.enable_only_active:
        G.silence_duration = 0
        G.notification_sent = False
        if G.event_logging:
            print("Reset silence duration due to plugin state change")
    
    # Remove existing timer if any
    obs.timer_remove(event_loop)
    # Only add timer if plugin is enabled
    if G.plugin_enabled:
        # Reset output tracking state so recording start/stop will be detected fresh
        G.prev_output_active = False
        G.silence_duration = 0
        G.notification_sent = False
        obs.timer_add(event_loop, G.tick)
        if G.event_logging:
            print(f"Plugin enabled - monitoring '{G.mic_source_name}'")
            print(f"Check interval: {G.tick / 1000}s, Silence threshold: {G.silence_threshold}s")
            print(f"Windows notification: {G.enable_windows_notification}")