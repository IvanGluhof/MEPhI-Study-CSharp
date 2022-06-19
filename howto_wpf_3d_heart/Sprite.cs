using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Media.Media3D;

namespace howto_wpf_3d_heart
{
    public abstract class Sprite
    {
        // The sprite can set this to false when it should stop.
        public bool Enabled;

        // The sprite's start time.
        public DateTime StartTime = DateTime.Now;

        // The 3D models to transform.
        public Model3DCollection Models = new Model3DCollection();

        // Move the models.
        public abstract void Move(DateTime now);
    }
}
