using System;
using System.IO;
using System.Threading.Tasks;

namespace workschedule
{
    public class Log
    {
        public enum EventType
        {
            Error,
            Info
        }
        private readonly TextWriter _writer;

        public Log(TextWriter writer)
        {
            _writer = writer;
            _writer.WriteLine($"Begin log {DateTime.Now:G}");
        }

        

        public async void WriteAsync(EventType eventType, string first, params string[] s)
        {
            await WriteInternalAsync($"[{DateTime.Now:T}]", $"{eventType,-5:G} - {first}");
            foreach (var str in s) await WriteInternalAsync(str);
        }

        private async Task WriteInternalAsync(string head, string tail)
        {
            await _writer.WriteAsync($"{head,12} {tail}");
        }

        private async Task WriteInternalAsync(string s)
        {
            await WriteInternalAsync(string.Empty, s);
        }
    }
}