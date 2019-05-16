using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MSHTML;
using SHDocVw;

namespace ConsoleApplication1
{
    internal class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            ShellWindows allBrowsers = new ShellWindows();
            Console.WriteLine($"IE Processes: {allBrowsers.Count}");

            var nternetExplorers = allBrowsers
                .Cast<InternetExplorer>()
                .ToArray();

            var changingBrowsers =
            nternetExplorers
                .Select((ie, index) =>
                {
//                    return Task.Run(() =>
//                    {
//                        ie.Navigate("https://redmine.x-code.pl");
//                        HTMLDocument ieDocument = ie.Document;
//
//                        var scriptTag = (IHTMLScriptElement)ieDocument.createElement("script");
//
//                        scriptTag.type = "text/javascript";
//                        scriptTag.src =
//                            @"function test(str){ alert('you clicked' + str);}";
//
//                        var head = (IHTMLElement)((IHTMLElementCollection)ieDocument.all.tags("head")).item(null, 0);
//                        
//                        ((HTMLHeadElement) head).appendChild((IHTMLDOMNode) scriptTag);
//
//                        ieDocument.parentWindow.execScript("test('siema eniu');");
//                    });

                    return StartSTATask(() =>
                    {
                        ie.Navigate("https://google.pl");
                        HTMLDocument ieDocument = ie.Document;

                        var scriptTag = (IHTMLScriptElement)ieDocument.createElement("script");

                        scriptTag.type = "text/javascript";
                        scriptTag.src =
                            @"function test(str){ alert('you clicked' + str);}";

                        var head = (IHTMLElement)((IHTMLElementCollection)ieDocument.all.tags("head")).item(null, 0);
                        
                        ((HTMLHeadElement) head).appendChild((IHTMLDOMNode) scriptTag);

                        ieDocument.parentWindow.execScript("test('siema eniu');");
                    });
                })
                .ToArray();

            Task.WaitAll(changingBrowsers);
            
            foreach (var nternetExplorer in nternetExplorers)
            {
                nternetExplorer.Visible = true;
            }
            
            Console.WriteLine();
        }
        
        public static Task StartSTATask(Action func)
        {
            var tcs = new TaskCompletionSource<object>();
            var thread = new Thread(() =>
            {
                try
                {
                    func();
                    tcs.SetResult(null);
                }
                catch (Exception e)
                {
                    tcs.SetException(e);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }
    }
}