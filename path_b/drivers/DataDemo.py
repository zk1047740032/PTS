######################################################################################################
# @file DataDemo.py
# @copyright HighFinesse GmbH.
# @author Lovas Szilard <lovas@highfinesse.de>
# @date 2018.09.15
# @version 0.1
#
# Homepage: http://www.highfinesse.com/
#
# @brief Python language example for demonstrating usage of
# wlmData.dll Set/Get function calls.
# Tested with Python 3.7. 64-bit Python requires 64-bit wlmData.dll and
# 32-bit Python requires 32-bit wlmData.dll.
# For more information see ctypes module documentation:
# https://docs.python.org/3/library/ctypes.html
# and/or WLM manual.pdf
#
# Changelog:
# ----------
# 2018.09.15
# v0.1 - Initial release
#

import sys

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

# Read type, version, revision and build number
version_type = dll.GetWLMVersion(0)
version_ver = dll.GetWLMVersion(1)
version_rev = dll.GetWLMVersion(2)
version_build = dll.GetWLMVersion(3)
print(f'WLM version: [{version_type}.{version_ver}.{version_rev}.{version_build}]')

# Read frequency
frequency = dll.GetFrequency(0.0)
if frequency == wlmConst.ErrWlmMissing:
    status_string = 'WLM inactive'
elif frequency == wlmConst.ErrNoSignal:
    status_string = 'No signal'
elif frequency == wlmConst.ErrBadSignal:
    status_string = 'Bad signal'
elif frequency == wlmConst.ErrLowSignal:
    status_string = 'Low signal'
elif frequency == wlmConst.ErrBigSignal:
    status_string = 'High signal'
else:
    status_string = 'WLM is running'

print(f'Status: {status_string}')

# Read temperaure
temperature = dll.GetTemperature(0.0)
if temperature <= wlmConst.ErrTemperature:
    print('Temperature: not available')
else:
    print(f'Temperature: {temperature:.1f} °C')

# Read pressure
pressure = dll.GetPressure(0.0)
if pressure <= wlmConst.ErrTemperature:
    print('Pressure: not available')
else:
    print(f'Pressure: {pressure:.1f} mbar')

# Read exposure of CCD arrays
exposure = dll.GetExposure(0)
if exposure == wlmConst.ErrWlmMissing:
    print('Exposure: WLM not active')
elif exposure == wlmConst.ErrNotAvailable:
    print('Exposure: not available')
else:
    print(f'Exposure: {exposure} ms')

# Read Ch1 Exposure mode
exposure_mode = dll.GetExposureMode(False)
if exposure_mode == 1:
    print('Ch1 Exposure: automatic')
else:
    print('Ch1 Exposure: manual')

# Read pulse mode settings
pulse_mode = dll.GetPulseMode(0)
if pulse_mode == 0:
    pulse_mode_string = 'Continuous Wave (cw) laser'
elif pulse_mode == 1:
    pulse_mode_string = 'Standard / Single / internally triggered pulsed'
elif pulse_mode == 2:
    pulse_mode_string = 'Double / Externally triggered pulsed'
else:
    pulse_mode_string = 'Other'
print(f'Pulse mode: {pulse_mode_string}')

# Read precision (wide/fine)
precision = dll.GetWideMode(0)
if precision == 0:
    precision_string = 'fine'
elif precision == 1:
    precision_string = 'wide'
else:
    precision_string = 'function not available'
print(f'Precision: {precision_string}')

# Print out frequency
if frequency == wlmConst.ErrOutOfRange:
    print('Ch1 error: out of range')
elif frequency <= 0:
    print(f'Ch1 error code: {frequency}')
else:
    print(f'Ch1 frequency: {frequency} THz')
