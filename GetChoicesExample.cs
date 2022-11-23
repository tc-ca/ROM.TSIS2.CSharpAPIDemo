using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROM.TSIS2.CSharpAPIDemo
{
    public static class GetChoicesExample
    {
        public static List<ChoiceItem> GetChoices(OptionSetMetadata optionSetMetadata)
        {
            List<ChoiceItem> choices = new List<ChoiceItem>();

            foreach (var item in optionSetMetadata.Options)
            {
                choices.Add(new ChoiceItem { 
                    EnglishLabel = item.Label.LocalizedLabels[0].Label,
                    FrenchLabel = item.Label.LocalizedLabels[1].Label,
                    Value = item.Value
                });
            }

            return choices;
        }
    }
}
