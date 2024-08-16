// I lost the source for this binary a long time ago but this is what ILSpy tells me the executable consists of.
// PowerPointTransformer, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// PowerPointTransformer.PowerPointTransformer

using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

public static int Main(string[] args)
{
	if (args.Length < 1)
	{
		return 1;
	}
	string fullPath = Path.GetFullPath(args[0]);
	if (!File.Exists(fullPath))
	{
		return 2;
	}
	Application application = (Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("91493441-5A91-11CF-8700-00AA0060263B")));
	Presentation presentation = application.Presentations.Open(fullPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
	Slides slides = presentation.Slides;
	foreach (Slide item in slides)
	{
		item.SlideShowTransition.AdvanceOnClick = MsoTriState.msoFalse;
		item.SlideShowTransition.AdvanceTime = 6.5f;
		item.SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
	}
	presentation.SlideShowSettings.LoopUntilStopped = MsoTriState.msoTrue;
	presentation.SaveAs(fullPath + "-transformed.pptx");
	return 0;
}
