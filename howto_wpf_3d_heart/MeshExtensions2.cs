using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Media.Media3D;

namespace howto_wpf_3d_heart
{
    public static partial class MeshExtensions
    {
        // Add a box representing a wall to the mesh.
        public static void AddWall(this MeshGeometry3D mesh,
            double length, double thickness, params Point3D[] points)
        {
//            mesh.AddBox()
        }

        // Make a new mesh that is a reflection of this one across the X-Y plane.
        // Reverse the order of triangles to maintain outward orientation.
        public static MeshGeometry3D ReflectZ(this MeshGeometry3D mesh)
        {
            MeshGeometry3D new_mesh = new MeshGeometry3D();

            // Add the positions.
            int num_points = mesh.Positions.Count;
            for (int i = 0; i < num_points; i++)
            {
                new_mesh.Positions.Add(new Point3D(
                    mesh.Positions[i].X,
                    mesh.Positions[i].Y,
                    -mesh.Positions[i].Z));
            }

            // Add the triangles.
            int num_indices = mesh.TriangleIndices.Count;
            for (int i = 0; i < num_indices; i += 3)
            {
                new_mesh.TriangleIndices.Add(mesh.TriangleIndices[i + 2]);
                new_mesh.TriangleIndices.Add(mesh.TriangleIndices[i + 1]);
                new_mesh.TriangleIndices.Add(mesh.TriangleIndices[i + 0]);
            }

            return new_mesh;
        }
    }
}