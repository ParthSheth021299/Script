Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Keyboard {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    
    public const int KEYEVENTF_KEYUP = 0x0002;
    
    public static void PressKey(byte key) {
        keybd_event(key, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(50);
        keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
    public static void PressCombo(byte modifier, byte key) {
        keybd_event(modifier, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(50);
        keybd_event(key, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(50);
        keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(modifier, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
"@

# Application "SSW" is automatically opened.
# We give it a few seconds to ensure it is visible and registered on the taskbar.
Start-Sleep -Seconds 6

# 2. Focus the Taskbar directly (Win + T)
[Keyboard]::PressCombo(0x5B, 0x54) 
Start-Sleep -Seconds 1

# 3. Press 'End' to jump to the very last application on the taskbar (the one we just launched)
[Keyboard]::PressKey(0x23)
Start-Sleep -Milliseconds 500

# 4. Open the jump list menu using the Menu/Apps Key
[Keyboard]::PressKey(0x5D)
Start-Sleep -Seconds 1

# 5. Detect OS Language to type the correct first letter!
$uiLang = (Get-UICulture).TwoLetterISOLanguageName
$keyToPress = 0x50 # Default 'P' (English: Pin to taskbar)

switch ($uiLang) {
    'fr' { $keyToPress = 0x45 } # 'E' (French: Épingler à la barre des tâches)
    'es' { $keyToPress = 0x41 } # 'A' (Spanish: Anclar a la barra de tareas)
    'de' { $keyToPress = 0x41 } # 'A' (German: An Taskleiste anheften)
    'it' { $keyToPress = 0x41 } # 'A' (Italian: Aggiungi alla barra delle applicazioni)
    'pt' { $keyToPress = 0x46 } # 'F' (Portuguese: Fixar na barra de tarefas)
}

[Keyboard]::PressKey($keyToPress)
Start-Sleep -Milliseconds 500

# 6. Press Enter to confirm the pin natively
[Keyboard]::PressKey(0x0D)
Start-Sleep -Seconds 1

# 7. Press Escape to clear the jump list if it stayed open
[Keyboard]::PressKey(0x1B)
