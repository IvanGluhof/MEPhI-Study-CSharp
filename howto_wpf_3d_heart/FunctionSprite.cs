using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Media.Media3D;

namespace howto_wpf_3d_heart
{
    public class FunctionSprite : Sprite
    {
        // The axis of rotation.
        public Vector3D RotationAxis;

        // Functions to generate the transformations.
        // Set these to null to skip a transformation.
        public Func<double, double> FAngle = null;
        public Func<double, double> FScaleX = null;
        public Func<double, double> FScaleY = null;
        public Func<double, double> FScaleZ = null;
        public Func<double, double> FOffsetX = null;
        public Func<double, double> FOffsetY = null;
        public Func<double, double> FOffsetZ = null;

        // Transform the models.
        public override void Move(DateTime now)
        {
            // If we're disabled, do nothing.
            if (!Enabled) return;

            // Get the elapsed time.
            double cumulative_seconds = (now - StartTime).TotalSeconds;

            // Create the necessary transformations.
            Transform3DGroup transforms = new Transform3DGroup();

            // Rotate.
            if (FAngle != null)
            {
                AxisAngleRotation3D rotation = new AxisAngleRotation3D(
                    RotationAxis, FAngle(cumulative_seconds));
                RotateTransform3D rotation_transform =
                    new RotateTransform3D(rotation);
                transforms.Children.Add(rotation_transform);
            }

            // Scale.
            if (FScaleX != null)
            {
                transforms.Children.Add(new ScaleTransform3D(
                    FScaleX(cumulative_seconds),
                    FScaleY(cumulative_seconds),
                    FScaleZ(cumulative_seconds)));
            }

            // Translate.
            if (FOffsetX != null)
            {
                transforms.Children.Add(new TranslateTransform3D(
                    FOffsetX(cumulative_seconds),
                    FOffsetY(cumulative_seconds),
                    FOffsetZ(cumulative_seconds)));
            }

            // Transform the models.
            foreach (GeometryModel3D model in Models)
                model.Transform = transforms;
        }
    }
}
