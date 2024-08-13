using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CFDiExcel.Test
{
    public class _01_Ingreso
    {
        [Fact]
        public void GenerarComprobante()
        {
            var xmlData = TestFileUtils.ReadFileAsString("01-Ingreso.xml");

            var comprobante = Program.DeserializeCfdi(xmlData);

            Assert.Equal("ESCUELA KEMPER URGATE", comprobante.Emisor.Nombre);
            Assert.Equal("2024-04-29T00:00:55", comprobante.Fecha);
            Assert.Equal("199.96", comprobante.Total);
        }
    }
}
