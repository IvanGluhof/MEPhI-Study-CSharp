//#define ShowWireframe
//#define ShowDetailedWireframe
#define SmoothTriangles
#define SpriteEnabled

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Windows.Media.Media3D;
using System.Windows.Threading;

namespace howto_wpf_3d_heart
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // The number of 100 ns timer ticks per second.
        private const long ns_per_second = 1000000000;
        private const long ticks_per_second = ns_per_second / 100;

        // The timer to manage animation.
        private DispatcherTimer AnimationTimer;

        // Sprites.
        private List<Sprite> Sprites = new List<Sprite>();

        // The heart model.
        private GeometryModel3D HeartModel;

        // The main object model group.
        private Model3DGroup MainModelGroup = new Model3DGroup();

        // Lights.
        private List<Light> Lights = new List<Light>();

        // The models.
        private GeometryModel3D AxesModel;
        private List<GeometryModel3D> AxesModels = new List<GeometryModel3D>();
        private List<GeometryModel3D> SolidModels = new List<GeometryModel3D>();

        // The camera.
        private PerspectiveCamera TheCamera;

        // The camera's current location.
        private double CameraPhi, CameraTheta, CameraR;

        // The amount by which we rotate during a 1 pixel manipulation.
        private double RotateRadiansPerPixel;

        // Acceleration for manipulation.
        private const double MovementAcceleration = 2;

        // Create the scene.
        // MainViewport is the Viewport3D defined
        // in the XAML code that displays everything.
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Make the border turn mouse events into touch events.
            MouseTouchDevice.RegisterEvents(borViewport);

            // Give the camera its initial position.
            TheCamera = new PerspectiveCamera();
            TheCamera.FieldOfView = 60;
            MainViewport.Camera = TheCamera;
            ResetCamera();

            // Define lights.
            DefineLights();

            // Define the models.
            DefineModels();

            // Add the group of models to a ModelVisual3D.
            ModelVisual3D model_visual = new ModelVisual3D();
            model_visual.Content = MainModelGroup;

            // Display the main visual to the viewport.
            MainViewport.Children.Add(model_visual);

            // Display the selected models.
            DisplaySelectedModels();

            // Make a sprite to spin the heart.
            TransformSprite sprite = new TransformSprite();
            sprite.Transforms = new Transform3DGroup();
            const double degrees_per_second = 360.0 / 4;
            Rotation3D rotation = new AxisAngleRotation3D(
                new Vector3D(0, 1, 0), degrees_per_second);
            sprite.Transforms.Children.Add(new RotateTransform3D(rotation));
            sprite.Models.Add(HeartModel);
#if SpriteEnabled
            sprite.Enabled = true;
#endif
            Sprites.Add(sprite);

            // Start the animation timer.
            AnimationTimer = new DispatcherTimer();
            AnimationTimer.Tick += AnimationTimer_Tick;
            double frames_per_second = 20;
            AnimationTimer.Interval = new TimeSpan((long)(1.0 / frames_per_second * ticks_per_second));
            AnimationTimer.Start();
        }

        // Move sprites that need to be moved.
        private bool TimeInitialized = false;
        private DateTime StartTime, LastTickTime;
        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (!TimeInitialized)
            {
                TimeInitialized = true;
                StartTime = DateTime.Now;
                LastTickTime = StartTime;
            }

            DateTime now = DateTime.Now;
            double cumulative_seconds = (now - StartTime).TotalSeconds;
            double elapsed_seconds = (now - LastTickTime).TotalSeconds;
            foreach (Sprite sprite in Sprites)
                if (sprite.Enabled) sprite.Move(now);

            LastTickTime = now;
        }

        // Define the lights.
        private void DefineLights()
        {
            Color color64 = Color.FromArgb(255, 128, 128, 128);
            Color color128 = Color.FromArgb(255, 255, 255, 255);
            Lights.Add(new AmbientLight(color64));
            Lights.Add(new DirectionalLight(color64,
                new Vector3D(-1.0, -1.0, -1.0)));
            Lights.Add(new DirectionalLight(color64,
                new Vector3D(1.0, 1.0, 1.0)));
        }

        // Create the models.
        private void DefineModels()
        {
            Cursor = Cursors.Wait;

            // Line thickness.
            const double line_thickness = 0.02;

            // Make the axes model.
            MeshGeometry3D axes_mesh = new MeshGeometry3D();
            axes_mesh.AddAxes(20.25, line_thickness, false, true, 2, 8);
            AxesModel = axes_mesh.MakeModel(Colors.Red);
            AxesModels.Add(AxesModel);

            // Make the heart.
            MeshGeometry3D heart_mesh = MakeHeartMesh();
#if ShowWireframe
            heart_mesh = heart_mesh.ToWireframe(0.05);
#elif ShowDetailedWireframe
            heart_mesh = heart_mesh.ToWireframe(0.01);
#endif
            HeartModel = heart_mesh.MakeModel(Colors.Red);
            SolidModels.Add(HeartModel);

            // Draw boxes along the axes for orientation.
            MeshGeometry3D x_mesh = new MeshGeometry3D();
            x_mesh.AddBox(new Point3D(5, 0, 0), 0.3, 0.3, 0.3);
            GeometryModel3D x_model = x_mesh.MakeModel(Colors.Red);
            AxesModels.Add(x_model);

            MeshGeometry3D y_mesh = new MeshGeometry3D();
            y_mesh.AddBox(new Point3D(0, 2, 0), 0.3, 0.3, 0.3);
            GeometryModel3D y_model = y_mesh.MakeModel(Colors.Green);
            AxesModels.Add(y_model);

            MeshGeometry3D z_mesh = new MeshGeometry3D();
            z_mesh.AddBox(new Point3D(0, 0, 4), 0.3, 0.3, 0.3);
            GeometryModel3D z_model = z_mesh.MakeModel(Colors.Blue);
            AxesModels.Add(z_model);

            Cursor = null;
        }

        // Make the heart mesh.
        private MeshGeometry3D MakeHeartMesh()
        {
            MeshGeometry3D mesh = new MeshGeometry3D();

            // Generate points along the edges of the heart in the X-Y plane.
            // Generate points counterclockwise.

            // Make an array of scales.
#if ShowWireframe || ShowDetailedWireframe || !SmoothTriangles
            const int num_slices = 5;
#else
            const int num_slices = 10;
#endif
            double[] scales = new double[num_slices];
            double dtheta = Math.PI / 2.0 / (num_slices + 1);
            for (int i = 0; i < num_slices; i++)
                scales[i] = Math.Cos(i * dtheta) / 4.0;

            // Make the array of lists to hold the polygons.
            List<Point3D>[] polygons = new List<Point3D>[num_slices];

            // Make the polygons.
#if ShowWireframe || ShowDetailedWireframe || !SmoothTriangles
            const int num_points = 10;
#else
            const int num_points = 100;
#endif
            double dt = 2 * Math.PI / num_points;
            double t = 0;
            for (int slice = 0; slice < num_slices; slice++)
            {
                polygons[slice] = new List<Point3D>();
                for (int i = 0; i < num_points; i++, t -= dt)
                    polygons[slice].Add(new Point3D(
                        X(t) * scales[slice],
                        Y(t) * scales[slice],
                        0));
            }

            // Set Z coordinates.
            // Loop through all polygons except the first.
            for (int slice = 1; slice < num_slices; slice++)
            {
                // Loop through the polygon's points.
                for (int i = 0; i < num_points; i++)
                {
                    // Get the distance from the point to the first polygon.
                    Point3D point = polygons[slice][i];
                    double dist = DistancePointToPolygon(point, polygons[0]);

                    // Option 1. Arc of a circle.
                    const double max_dist = 1.5;
                    if (dist > max_dist) dist = max_dist;

                    // Option 1. Curve along an arc of a circle.
                    //point.Z = Math.Sqrt(max_dist * max_dist - (dist - max_dist) * (dist - max_dist));

                    // Option 2. Sqrt curve.
                    point.Z = Math.Sqrt(dist);

                    // Update the point.
                    polygons[slice][i] = point;
                }
            }

            // Connect slices.
#if ShowWireframe
            for (int slice = 0; slice < num_slices; slice++)
                mesh.AddPolygonWireframe(0.05, polygons[slice].ToArray());
#else
            for (int slice = 0; slice < num_slices - 1; slice++)
                ConnectSlices(mesh, polygons[slice], polygons[slice + 1]);
#endif

            // Add the final polygon on top.
#if ShowWireframe
            mesh.AddPolygonWireframe(0.05, polygons[num_slices - 1].ToArray());
#else
            mesh.AddPolygon(polygons[num_slices - 1].ToArray());

            // Make a reflection across the X-Y plane.
            MeshGeometry3D new_mesh = mesh.ReflectZ();

            // Merge the reflection into the original mesh.
            mesh.MergeWith(new_mesh);
#endif

#if SmoothTriangles
            // Smooth if desired.
            mesh.Smooth();
#endif

            return mesh;
        }

        // Make triangles connecting the two polygons point-for-point.
        private void ConnectSlices(MeshGeometry3D mesh, List<Point3D> polygon1, List<Point3D> polygon2)
        {
            int num_points = polygon1.Count;
            Point3D last_point1 = polygon1[num_points - 1];
            Point3D last_point2 = polygon2[num_points - 1];
            for (int i = 0; i < num_points; i++)
            {
                Point3D next_point1 = polygon1[i];
                Point3D next_point2 = polygon2[i];

                // Triangulate the "rectangle" between the four points.
                TriangulateRectangle(mesh, last_point1, next_point1, next_point2, last_point2);

                // Get ready for the next "rectangle."
                last_point1 = next_point1;
                last_point2 = next_point2;
            }
        }

        // The curve's parametric equations.
        private double Y(double t)
        {
            return
                13 * Math.Cos(t) -
                5 * Math.Cos(2 * t) -
                2 * Math.Cos(3 * t) -
                Math.Cos(4 * t);
        }
        private double X(double t)
        {
            double sin_t = Math.Sin(t);
            return 16 * sin_t * sin_t * sin_t;
        }

        // Make four triangle to represent this "rectangle."
        // The points must be outward oriented.
        // Do nothing unless all four points have Z >= 0.
        private void TriangulateRectangle(MeshGeometry3D mesh,
            Point3D p00, Point3D p01, Point3D p11, Point3D p10)
        {
            // Find the average of the points.
            Point3D middle = new Point3D(
                (p00.X + p01.X + p11.X + p10.X) * 0.25,
                (p00.Y + p01.Y + p11.Y + p10.Y) * 0.25,
                (p00.Z + p01.Z + p11.Z + p10.Z) * 0.25);

            // Make four triangles.
#if SmoothTriangles
            mesh.AddSmoothTriangle(p00, p01, middle);
            mesh.AddSmoothTriangle(p01, p11, middle);
            mesh.AddSmoothTriangle(p11, p10, middle);
            mesh.AddSmoothTriangle(p10, p00, middle);
#else
            mesh.AddTriangle(p00, p01, middle);
            mesh.AddTriangle(p01, p11, middle);
            mesh.AddTriangle(p11, p10, middle);
            mesh.AddTriangle(p10, p00, middle);
#endif
        }

#region FunctionSprite functions

        private const double TwoPi = 2.0 * Math.PI;
        // 1 rotation per second.
        private double FAngle(double cumulative_seconds)
        {
            return cumulative_seconds * 360.0;
        }

        // Scale between 1 and 2 every 2 seconds.
        private double FScaleX(double cumulative_seconds)
        {
            return 1.5 + 0.5 * Math.Sin(TwoPi * cumulative_seconds / 2.0);
        }

        // Make an ellipse in the XZ plane every 4 seconds.
        private double FOffsetX(double cumulative_seconds)
        {
            return 3 * Math.Cos(TwoPi * cumulative_seconds / 4.0);
        }
        private double FOffsetY(double cumulative_seconds)
        {
            return 0;
        }

        private double FOffsetZ(double cumulative_seconds)
        {
            return 5 * Math.Sin(TwoPi * cumulative_seconds / 4.0);
        }

#endregion FunctionSprite functions

        // Position the camera.
        private void PositionCamera()
        {
            // Calculate the camera's position in Cartesian coordinates.
            double y = CameraR * Math.Sin(CameraPhi);
            double hyp = CameraR * Math.Cos(CameraPhi);
            double x = hyp * Math.Cos(CameraTheta);
            double z = hyp * Math.Sin(CameraTheta);
            TheCamera.Position = new Point3D(x, y, z);

            // Look toward the origin.
            TheCamera.LookDirection = new Vector3D(-x, -y, -z);
            TheCamera.Transform =
                new TranslateTransform3D(0, -1, 0);

            // Set the Up direction.
            TheCamera.UpDirection = new Vector3D(0, 1, 0);
        }

        private void chkItem_Click(object sender, RoutedEventArgs e)
        {
            DisplaySelectedModels();
        }

        // Get ready to process a manipulation.
        private void Border_ManipulationStarting(object sender, ManipulationStartingEventArgs e)
        {
            // No need for this event to continue.
            e.Handled = true;

            // Allow translation and scaling.
            e.Mode = ManipulationModes.All;

            // Calculate rotation in degrees per pixel
            double field_width_pixels = MainViewport.ActualWidth;
            double field_radians = TheCamera.FieldOfView * Math.PI / 180.0;
            RotateRadiansPerPixel = field_radians / field_width_pixels * MovementAcceleration;
        }

        // Process rotation and scaling manipulations.
        private void Border_ManipulationDelta(object sender, ManipulationDeltaEventArgs e)
        {
            // Rotate horizontally.
            CameraTheta += e.DeltaManipulation.Translation.X * RotateRadiansPerPixel;

            // Rotate vertically.
            CameraPhi += e.DeltaManipulation.Translation.Y * RotateRadiansPerPixel;
            if (CameraPhi > Math.PI / 2) CameraPhi = Math.PI / 2;
            if (CameraPhi < -Math.PI / 2) CameraPhi = -Math.PI / 2;

            // Zoom in or out.
            double scale = 2 - Math.Max(
                e.DeltaManipulation.Scale.X,
                e.DeltaManipulation.Scale.Y);
            if (scale > 0)
            {
                CameraR *= scale;
                lblDistance.Content = CameraR.ToString("0.0");
            }

            PositionCamera();
        }

        // Prepare for inertia.
        private void Border_ManipulationInertiaStarting(object sender, ManipulationInertiaStartingEventArgs e)
        {
            // No need for this event to continue.
            e.Handled = true;

            // Translation (which is translated into rotation for CameraTheta and CameraPhi).
            // Decelerate by 5" per second each second.
            // (This is somewhat arbitrary.)
            e.TranslationBehavior.DesiredDeceleration = 5.0 * 96.0 / (1000.0 * 1000.0);

            // Expansion (which is translated into zooming in or out with CameraR).
            // 1 inch per second each second.
            // This seems to work when zooming in but not out.
            e.ExpansionBehavior.DesiredDeceleration = 1 * 96.0 / (1000.0 * 1000.0);

            // Skip rotation.
        }

        // Reset the camera.
        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            ResetCamera();
        }

        private void ResetCamera()
        {
            CameraTheta = 1;
            CameraPhi = 0.3;
            CameraR = 12.0;
            lblDistance.Content = CameraR;
            PositionCamera();
        }

        // Display the selected models.
        private void DisplaySelectedModels()
        {
            MainModelGroup.Children.Clear();

            foreach (Light light in Lights)
                MainModelGroup.Children.Add(light);

            if (chkAxes.IsChecked.Value)
            {
                foreach (Model3D model in AxesModels)
                    MainModelGroup.Children.Add(model);
            }

            foreach (GeometryModel3D model in SolidModels)
                MainModelGroup.Children.Add(model);
        }

        // Display the point clicked.
        private void MainViewport_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // Get the mouse's position relative to the viewport.
            Point mouse_pos = e.GetPosition(MainViewport);

            // Perform the hit test.
            HitTestResult result =
                VisualTreeHelper.HitTest(MainViewport, mouse_pos);

            // Display information about the hit.
            RayMeshGeometry3DHitTestResult mesh_result =
                result as RayMeshGeometry3DHitTestResult;
            if (mesh_result != null)
            {
                //Console.WriteLine("Point hit: (" +
                //    mesh_result.PointHit.ToString() + ")");
                Console.WriteLine("new Point3D(" +
                    mesh_result.PointHit.X + ", " +
                    mesh_result.PointHit.Y + ", " +
                    mesh_result.PointHit.Z + "),");

                //MeshGeometry3D mesh = new MeshGeometry3D();
                //Vector3D offset = mesh_result.PointHit - new Point3D(0, 0, 0);
                //mesh.AddIcosahedron(0.12, new TranslateTransform3D(offset));
                //GeometryModel3D model = mesh.MakeModel(Colors.out);
                //SolidModels.Add(model);
                //DisplaySelectedModels();
            }
        }

#region Geometry

        // Return True if the point is in the polygon ignoring Z coordinates.
        public bool PointIsInPolygon(Point3D point, List<Point3D> points)
        {
            // Get the angle between the point and all pairs of vertices.
            int num_points = points.Count;
            double total_angle = 0;
            Point3D last_point = points[num_points - 1];
            for (int i = 0; i < num_points; i++)
            {
                Point3D next_point = points[i];
                total_angle += GetAngle(last_point, point, next_point);
                last_point = next_point;
            }

            // The total angle should be 2 * PI or -2 * PI if
            // the point is in the polygon.
            // It should be close to zero
            // if the point is outside the polygon.
            return (Math.Abs(total_angle) > 0.000001);
        }

        // Return the angle ABC ignoring Z coordinates.
        // Return a value between PI and -PI.
        public static double GetAngle(Point3D A, Point3D B, Point3D C)
        {
            // Get the dot product.
            double dot_product = DotProductXZ(A, B, C);

            // Get the cross product.
            double cross_product_length = CrossProductLengthXZ(A, B, C);

            // Calculate the angle.
            return Math.Atan2(cross_product_length, dot_product);
        }

        // Return the dot product AB · BC ignoring Z coordinates.
        // Note that AB · BC = |AB| * |BC| * Cos(theta).
        private static double DotProductXZ(Point3D A, Point3D B, Point3D C)
        {
            // Get the vectors' coordinates.
            double BAx = A.X - B.X;
            double BAy = A.Y - B.Y;
            double BCx = C.X - B.X;
            double BCy = C.Y - B.Y;
            return BAx * BCx + BAy * BCy;
        }

        // Return the length of the cross product AB x BC ignoring Z coordinates.
        // Note that AB x BC = |AB| * |BC| * Sin(theta).
        private static double CrossProductLengthXZ(Point3D A, Point3D B, Point3D C)
        {
            // Get the vectors' coordinates.
            double BAx = A.X - B.X;
            double BAy = A.Y - B.Y;
            double BCx = C.X - B.X;
            double BCy = C.Y - B.Y;

            // Calculate the Y component of the cross product.
            return (BAx * BCy - BAy * BCx);
        }

        // Return the minimum distance from
        // the point to a vertex on the polygon.
        private double DistancePointToPolygonVertex(Point3D point, List<Point3D> polygon)
        {
            double best_dist = double.MaxValue;
            foreach (Point3D vertex in polygon)
            {
                double test_dist = (point - vertex).Length;
                if (best_dist > test_dist) best_dist = test_dist;
            }
            return best_dist;
        }

        // Return the minimum distance in the X-Y plane
        // from the point to a segment on the polygon.
        private double DistancePointToPolygon(Point3D point, List<Point3D> polygon)
        {
            Point3D closest;
            double best_dist = double.MaxValue;
            Point3D last_vertex = polygon[polygon.Count - 1];
            foreach (Point3D vertex in polygon)
            {
                double test_dist = FindDistanceToSegment(point, last_vertex, vertex, out closest);
                if (best_dist > test_dist) best_dist = test_dist;
                last_vertex = vertex;
            }
            return best_dist;
        }

        // Calculate the distance between point pt and
        // the segment p1 --> p2 in the X-Y plane.
        private double FindDistanceToSegment(
            Point3D pt, Point3D p1, Point3D p2, out Point3D closest)
        {
            const double tiny = 0.00001;
            if ((p1 - p2).Length < tiny)
            {
                // It's a point not a line segment.
                closest = p1;
            }
            else
            {
                double dx = p2.X - p1.X;
                double dy = p2.Y - p1.Y;

                // Calculate the t that minimizes the distance.
                double t = ((pt.X - p1.X) * dx + (pt.Y - p1.Y) * dy) / (dx * dx + dy * dy);

                // See if this represents one of the segment's
                // end points or a point in the middle.
                if (t < 0) closest = p1;
                else if (t > 1) closest = p2;
                else
                {
                    double dz = p2.Z - p1.Z;
                    closest = new Point3D(
                        p1.X + t * dx,
                        p1.Y + t * dy,
                        p1.Z + t * dz);
                }
            }

            // Return the distance (ignoring the Z coordinate) to the closest point.
            double xdiff = pt.X - closest.X;
            double ydiff = pt.Y - closest.Y;
            return Math.Sqrt(xdiff * xdiff + ydiff * ydiff);
        }

        // Find the point of intersection between the lines
        // p1 --> p2 and p3 --> p4 ignoring Z coordinates (in the X-Y plane).
        private void FindIntersection(
            Point3D p1, Point3D p2, Point3D p3, Point3D p4,
            out bool lines_intersect, out bool segments_intersect,
            out Point3D intersection,
            out Point3D close_p1, out Point3D close_p2,
            out double t1, out double t2)
        {
            // Get the segments' parameters.
            double dx12 = p2.X - p1.X;
            double dy12 = p2.Y - p1.Y;
            double dx34 = p4.X - p3.X;
            double dy34 = p4.Y - p3.Y;

            // Solve for t1 and t2
            double denominator = (dy12 * dx34 - dx12 * dy34);

            t1 = ((p1.X - p3.X) * dy34 + (p3.Y - p1.Y) * dx34) / denominator;
            if (double.IsInfinity(t1))
            {
                // The lines are parallel (or close enough to it).
                lines_intersect = false;
                segments_intersect = false;
                Point3D infinity = new Point3D(double.NaN, double.NaN, double.NaN);
                intersection = infinity;
                close_p1 = infinity;
                close_p2 = infinity;
                t1 = double.NaN;
                t2 = double.NaN;
                return;
            }
            lines_intersect = true;

            t2 = ((p3.X - p1.X) * dy12 + (p1.Y - p3.Y) * dx12) / -denominator;

            // Find the point of intersection.
            double dz12 = p2.Z - p1.Z;
            intersection = new Point3D(p1.X + dx12 * t1, p1.Y + dy12 * t1, p1.Z + dz12 * t1);

            // The segments intersect if t1 and t2 are between 0 and 1.
            segments_intersect =
                ((t1 >= 0) && (t1 <= 1) &&
                 (t2 >= 0) && (t2 <= 1));

            // Find the closest points on the segments.
            if (t1 < 0) t1 = 0;
            else if (t1 > 1) t1 = 1;

            if (t2 < 0) t2 = 0;
            else if (t2 > 1) t2 = 1;

            close_p1 = new Point3D(p1.X + dx12 * t1, p1.Y + dy12 * t1, p1.Z + dy12 * t1);
            double dz34 = p4.Z - p3.Z;
            close_p2 = new Point3D(p3.X + dx34 * t2, p3.Y + dy34 * t2, p3.Z + dz34 * t2);
        }

        // Return true if the segment intersects the polygon.
        private bool SegmentIntersectsPolygon(Point3D point1, Point3D point2, List<Point3D> polygon)
        {
            // See where the segment intersects the polygon (if it does).
            bool segments_intersect;
            Point3D intersection, close_p1, close_p2;
            double t1, t2;
            int segment_number;

            FindSegmentPolygonIntersection(point1, point2, polygon,
                out segments_intersect,
                out intersection,
                out close_p1, out close_p2,
                out t1, out t2, out segment_number);

            // Count it if the intersection is not at the end of a segment.
            const double tiny = 0.01;
            return ((t1 > tiny) && (t1 < 1 - tiny) &&
                    (t2 > tiny) && (t2 < 1 - tiny));
        }

        // Return the point where the segment intersects the polygon (if it does).
        private void FindSegmentPolygonIntersection(Point3D point1, Point3D point2, List<Point3D> polygon,
            out bool segments_intersect,
            out Point3D intersection,
            out Point3D close_p1, out Point3D close_p2,
            out double t1, out double t2,
            out int segment_number)
        {
            // Initialize output variables.
            Point3D no_point = new Point3D(double.NaN, double.NaN, double.NaN);
            segments_intersect = false;
            intersection = no_point;
            close_p1 = no_point;
            close_p2 = no_point;
            t1 = double.NaN;
            t2 = double.NaN;
            segment_number = int.MinValue;

            bool lines_intersect;

            // See if the segment intersects any of the polygon's edges.
            int num_points = polygon.Count;
            Point3D last_point = polygon[num_points - 1];
            for (int i = 0; i < num_points; i++)
            {
                Point3D next_point = polygon[i];
                FindIntersection(point1, point2, last_point, next_point,
                    out lines_intersect, out segments_intersect,
                    out intersection, out close_p1, out close_p2,
                    out t1, out t2);

                // Don't count it if the intersection is at the end of a segment.
                if (segments_intersect)
                {
                    segment_number = i - 1;
                    return;
                }
                last_point = next_point;
            }

            // The segment doesn't intersect the polygon.
            // Initialize output variables.
            lines_intersect = false;
            segments_intersect = false;
            intersection = no_point;
            close_p1 = no_point;
            close_p2 = no_point;
            t1 = double.NaN;
            t2 = double.NaN;
            segment_number = int.MinValue;
        }

#endregion Geometry

    }
}
