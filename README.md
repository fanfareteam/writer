# Project Writer 
Project Writer is a lightweight, state-aware word processor built for the .NET 4.0 framework. It bridges the gap between the iconic design languages of the late 2000s and the modern era, offering a seamless experience from Windows XP to Windows 11.

# Features
## Multi-Era Theming Engine
Project Writer detects your operating system and adjusts its design language accordingly:
- Luna (Windows XP): Deep royal blue gradients with iconic green "Start-style" highlights.
- Aero (Windows Vista/7): Real DWM Glass transparency with "Golden Glow" button states.
- Classic (2009): The refined "Office-style" blue gradient look.
- UWP (Windows 10/11): Clean, flat design with a full Dark Mode implementation.
## Advanced Editor Capabilities
- A4 Canvas: A dedicated 794x1123 pixel workspace mimicking standard paper dimensions.
- Live Font Rendering: The font selection dropdown renders typefaces in their actual style using GDI+ OwnerDraw.
- Zoom Engine: Smooth scaling from 50% to 600%.
- Page Tracking: Real-time page counting based on line-density and vertical scroll position.
## Global Support
- Auto-Localization: Automatically detects System Language (supports English and German Beta).
- File Associations: Built-in support for "Open With" via the Windows Shell.
- Smart Recent Files: Intelligent history tracking stored in a lightweight app.dat configuration file.
# Technical Implementation
Requirements:
- Framework: .NET Framework 4.0
- IDE: Visual Studio 2010 or newer
- OS: Windows XP SP3, Vista SP2, 7, 8, 10, or 11
### State-Based Navigation
Project Writer uses a custom state engine to toggle between the Welcome Screen and the Editor UI. This prevents the "cluttered" feeling of traditional WinForms apps by only rendering the controls necessary for the current task.

# Installation & Usage
1. Download the latest update from the Releases tab.
2. Run the application. On first launch, you will be prompted for your preferred language.
3. To associate .rtf files, right-click a file > Open With > Choose another app and point it to writer.exe.

### Beta Disclaimer (.docx)
Support for .docx is currently in Beta. Project Writer focuses on high-fidelity RTF and TXT manipulation. If a Word document contains unsupported features (WordArt, Macros), the application will notify the user rather than risking a corrupt save.

## Credits:
Made by fanfare. Lead, developer and idea - juzadss, 88g (since 0.4.1) - design.
