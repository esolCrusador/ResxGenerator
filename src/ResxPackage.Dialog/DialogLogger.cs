using System.Collections.Generic;
using ResourcesAutogenerate;

namespace ResxPackage.Dialog
{
    public class DialogLogger : ILogger
    {
        private readonly List<string> _messagesList;

        public DialogLogger(List<string> messagesList)
        {
            _messagesList = messagesList;
        }

        public void Log(string message)
        {
            _messagesList.Add(message);
        }
    }
}
