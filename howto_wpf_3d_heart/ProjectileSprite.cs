using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Media.Media3D;

namespace howto_wpf_3d_heart
{
    public class ProjectileSprite : FunctionSprite
    {
        public double GravitationalConstant;
        public Point3D InitialLocation;
        public Vector3D InitialVelocity;
        public ProjectileSprite(Point3D initial_location, Vector3D initial_velocity, double gravitational_constant)
        {
            InitialLocation = initial_location;
            InitialVelocity = initial_velocity;
            GravitationalConstant = gravitational_constant;

            FOffsetX = FuncOffsetX;
            FOffsetY = FuncOffsetY;
            FOffsetZ = FuncOffsetZ;
        }

        // Calculate the object's current position.
        public double FuncOffsetX(double cumulative_seconds)
        {
            return InitialLocation.X +
                cumulative_seconds * InitialVelocity.X;
        }

        public double FuncOffsetY(double cumulative_seconds)
        {
            return InitialLocation.Y +
                cumulative_seconds * InitialVelocity.Y +
                GravitationalConstant / 2.0 * cumulative_seconds * cumulative_seconds;
        }

        public double FuncOffsetZ(double cumulative_seconds)
        {
            return InitialLocation.Z +
                cumulative_seconds * InitialVelocity.Z;
        }

        // Return the current position.
        public Point3D CurrentPosition(double cumulative_seconds)
        {
            return new Point3D(
                FuncOffsetX(cumulative_seconds),
                FuncOffsetY(cumulative_seconds),
                FuncOffsetZ(cumulative_seconds));
        }

        // Return the current velocity.
        public Vector3D CurrentVelocity(double cumulative_seconds)
        {
            return new Vector3D(
                InitialVelocity.X,
                InitialVelocity.Y - GravitationalConstant * cumulative_seconds,
                InitialVelocity.Z);
        }

        // Move.
        public override void Move(DateTime now)
        {
            base.Move(now);

            // See if we should stop.
            double cumulative_seconds = (now - StartTime).TotalSeconds;
            if (FuncOffsetY(cumulative_seconds) < 0) Enabled = false;
        }
    }
}
