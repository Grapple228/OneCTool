using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ITworks.Brom.Types;

namespace OneC.Interfaces
{
    internal interface IOneCActions
    {
        public Ссылка Link { get; set; }
        public void AddTo1C();
    }
}
