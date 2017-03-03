using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ClippyBot.Dialogs
{
    [Serializable]
    public enum OfficeOperation
    {
        Reply,
        ReplyAll,
        Chart,
        Range,
        Image,
        Paragraph
    }
}