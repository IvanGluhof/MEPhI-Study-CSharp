using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Media.Media3D;

namespace howto_wpf_3d_heart
{
    public class TransformSprite : Sprite
    {
        // The change transformations per second.
        public Transform3DGroup Transforms = new Transform3DGroup();

        // Transform the models.
        public override void Move(DateTime now)
        {
            // If we're disabled, do nothing.
            if (!Enabled) return;

            // Get the elapsed time.
            double cumulative_seconds = (now - StartTime).TotalSeconds;

            // Create the necessary transformations.
            Transform3DGroup transforms = new Transform3DGroup();

            foreach (Transform3D new_transforms in Transforms.Children)
            {
                if (new_transforms is RotateTransform3D)
                {
                    // Create a rotation transform with the angle scaled by cumulative elapsed seconds.
                    RotateTransform3D old_transform = new_transforms as RotateTransform3D;
                    if (old_transform.Rotation is AxisAngleRotation3D)
                    {
                        AxisAngleRotation3D aa_rotation =
                            old_transform.Rotation as AxisAngleRotation3D;
                        AxisAngleRotation3D rotation = new AxisAngleRotation3D(
                            aa_rotation.Axis, aa_rotation.Angle * cumulative_seconds);

                        RotateTransform3D new_transform = new RotateTransform3D(rotation);
                        transforms.Children.Add(new_transform);
                    }
                    else if (old_transform.Rotation is QuaternionRotation3D)
                    {
                        QuaternionRotation3D q_rotation =
                            old_transform.Rotation as QuaternionRotation3D;

                        Quaternion new_quaternion = new Quaternion(
                            q_rotation.Quaternion.Axis, q_rotation.Quaternion.Angle * cumulative_seconds);
                        QuaternionRotation3D new_rotation = new QuaternionRotation3D(new_quaternion);
                        RotateTransform3D new_transform = new RotateTransform3D(new_rotation);
                        transforms.Children.Add(new_transform);
                    }
                }
                else if (new_transforms is ScaleTransform3D)
                {
                    // Create a scale transform with the scale factors scaled by cumulative elapsed seconds.
                    ScaleTransform3D old_transform = new_transforms as ScaleTransform3D;
                    ScaleTransform3D new_transform = new ScaleTransform3D(
                        old_transform.ScaleX * cumulative_seconds,
                        old_transform.ScaleY * cumulative_seconds,
                        old_transform.ScaleZ * cumulative_seconds,
                        old_transform.CenterX,
                        old_transform.CenterY,
                        old_transform.CenterZ);
                    transforms.Children.Add(new_transform);
                }
                else if (new_transforms is TranslateTransform3D)
                {
                    TranslateTransform3D old_transform = new_transforms as TranslateTransform3D;
                    TranslateTransform3D new_transform = new TranslateTransform3D(
                        old_transform.OffsetX * cumulative_seconds,
                        old_transform.OffsetY * cumulative_seconds,
                        old_transform.OffsetZ * cumulative_seconds);
                    transforms.Children.Add(new_transform);
                }
            }

            // Transform the models.
            foreach (GeometryModel3D model in Models)
                model.Transform = transforms;
        }
    }
}