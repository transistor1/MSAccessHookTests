using System;
using System.Windows.Forms;
using System.Collections.Generic;

class Script
{
	[STAThread]
	static public void Main(string[] args)
	{
		List<string>[] lists= new List<string>[5];
        for (int i = 0; i < 5; i++)
        {
            lists[i].Add("abc");
        }

		for (int i = 0; i < args.Length; i++)
		{
			Console.WriteLine(args[i]);
		}
	}
}

