######################################################################################################
# @file PatternDemo.py
# @copyright HighFinesse GmbH.
# @version 0.1
#
# Homepage: http://www.highfinesse.com/
#

import sys
import ctypes

import matplotlib.pyplot as plt

# wlmData.dll related imports
import wlmConst
import wlmData

# Load wlmData library. If needed, adjust the path by passing it to LoadDLL()!
try:
    dll = wlmData.LoadDLL()
    #dll = wlmData.LoadDLL('/path/to/your/libwlmData.so')
except OSError as err:
    sys.exit(f'{err}\nPlease check if the wlmData DLL is installed correctly!')

# Check the number of WLM server instances
if dll.GetWLMCount(0) == 0:
    sys.exit('There is no running WLM server instance.')

# Enable pattern export
dll.SetPattern(wlmConst.cSignal1Interferometers, wlmConst.cPatternEnable)
two_patterns = dll.SetPattern(wlmConst.cSignal1WideInterferometer, wlmConst.cPatternEnable) \
    == wlmConst.ResERR_NoErr

# Request pattern parameters (these don't change later)
pattern_item_size_short = dll.GetPatternItemSize(wlmConst.cSignal1Interferometers)
pattern_item_count_short = dll.GetPatternItemCount(wlmConst.cSignal1Interferometers)

if pattern_item_size_short == 2:
    pattern_short = (ctypes.c_int16 * pattern_item_count_short)()
elif pattern_item_size_short == 4:
    pattern_short = (ctypes.c_int32 * pattern_item_count_short)()
else:
    sys.exit('Unknown pattern data format')

# More precise wavelength meters have an additional photodiode array:
if two_patterns:
    pattern_item_size_long = dll.GetPatternItemSize(wlmConst.cSignal1WideInterferometer)
    pattern_item_count_long = dll.GetPatternItemCount(wlmConst.cSignal1WideInterferometer)

    if pattern_item_size_long == 2:
        pattern_long = (ctypes.c_int16 * (pattern_item_count_long))()
    elif pattern_item_size_long == 4:
        pattern_long = (ctypes.c_int32 * pattern_item_count_long)()
    else:
        sys.exit('Unknown pattern data format')

# Set up Matplotlib
fig, axs = plt.subplots(2 if two_patterns else 1, 1)
if not two_patterns:
    axs = [axs,]

# Request pattern data.
# If synchronization with measurements is desired, please perform this with the
# callback or WaitForWLMEvent mechanism. If a multichannel switcher is attached,
# use GetPatternDataNum to distinguish different channels.
dll.GetPatternData(wlmConst.cSignal1Interferometers, pattern_short)
if two_patterns:
    dll.GetPatternData(wlmConst.cSignal1WideInterferometer, pattern_long)

# Plot using Matplotlib
axs[0].plot(pattern_short)
if two_patterns:
    axs[1].plot(pattern_long)
plt.show()
