using System.Windows.Controls;
using System.AddIn;
using System.AddIn.Pipeline;

using PrincetonInstruments.LightField.AddIns;

namespace LightFieldAddInSamples
{
    /// <summary>
    /// BAP Lab – Simple Multipoint Spectroscope
    /// Proof-of-concept: serial port scanning + camera acquisition + PNG save.
    ///
    /// NOTE: QualificationData("IsSample","True") is deliberately absent so that
    /// LightField does NOT group this add-in with the stock sample add-ins and it
    /// always appears at the top of the Add-in Manager list.
    /// </summary>
    [AddIn("BAP Lab – Simple Multipoint Spectroscope",
           Version = "1.0.0",
           Publisher = "BAP Lab",
           Description = "PoC: scan serial ports for Marlin stage, acquire images, save PNG.")]
    public class SimpleMultipointSpectroscopeAddin : AddInBase, ILightFieldAddIn
    {
        private BAP_Lab_SimpleMultipointSpectroscope.SimpleMultipointSpectroscopeControl control_;

        ///////////////////////////////////////////////////////////////////////////
        public UISupport UISupport
        {
            get { return UISupport.ExperimentView; }
        }

        ///////////////////////////////////////////////////////////////////////////
        public void Activate(ILightFieldApplication app)
        {
            LightFieldApplication = app;

            control_ = new BAP_Lab_SimpleMultipointSpectroscope.SimpleMultipointSpectroscopeControl(app);

            // Wrap in a ScrollViewer so its contents are reachable on small screens
            ScrollViewer sv = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
                Content                        = control_
            };

            ExperimentViewElement = sv;
        }

        ///////////////////////////////////////////////////////////////////////////
        public void Deactivate()
        {
            // Nothing to dispose in this PoC – serial port is never opened
        }

        ///////////////////////////////////////////////////////////////////////////
        public override string UIExperimentViewTitle
        {
            get { return "BAP – Multipoint Spectroscope"; }
        }
    }
}
