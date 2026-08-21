using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelCalibrationAddin.Core.Tests
{
    [TestClass]
    public sealed class ShortcutConfigurationTests
    {
        [TestMethod]
        public void DefaultGenerationShortcutIsSingleFunctionKey()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration());

            Assert.AreEqual("F6", normalized.GenerateShortcutKey);
        }

        [TestMethod]
        public void ShortcutConfigurationNormalizesFunctionKeyCase()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration
            {
                GenerateShortcutKey = " f9 "
            });

            Assert.AreEqual("F9", normalized.GenerateShortcutKey);
        }

        [TestMethod]
        public void EmptyShortcutDisablesKeyboardTrigger()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration
            {
                GenerateShortcutKey = string.Empty
            });

            Assert.AreEqual(string.Empty, normalized.GenerateShortcutKey);
        }

        [TestMethod]
        public void UnsupportedShortcutFallsBackToDefault()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration
            {
                GenerateShortcutKey = "A"
            });

            Assert.AreEqual("F6", normalized.GenerateShortcutKey);
        }

        [TestMethod]
        public void ModifierShortcutIsCanonicalized()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration
            {
                GenerateShortcutKey = "shift+ctrl+f6"
            });

            Assert.AreEqual("Ctrl+Shift+F6", normalized.GenerateShortcutKey);
        }

        [TestMethod]
        public void ModifierShortcutCanUseRegularKey()
        {
            var normalized = new GenerationConfigurationStore().Normalize(new GenerationConfiguration
            {
                GenerateShortcutKey = "ctrl+r"
            });

            Assert.AreEqual("Ctrl+R", normalized.GenerateShortcutKey);
        }
    }
}
