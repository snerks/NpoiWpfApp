using System;
using System.Collections.Generic;
using System.Text;

namespace NpoiWpfApp.Models
{
    internal class SampleClass
    {
        public string Header1 { get; set; }
        public string Header2 { get; set; }
        public bool IsHeader2Null => Header2 == null;

        public bool? IsEnabledRawValue { get; set; }

        public bool IsEnabledRawValueIsNull => IsEnabledRawValue == null;

        public bool IsEnabled => IsEnabledRawValue == true;
    }
}
