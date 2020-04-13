using Rhino;
using Rhino.Geometry;
using Rhino.DocObjects;
using Rhino.Collections;

using GH_IO;
using GH_IO.Serialization;
using Grasshopper;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;

using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;



/// <summary>
/// This class will be instantiated on demand by the Script component.
/// </summary>
public class Script_Instance : GH_ScriptInstance
{
#region Utility functions
  /// <summary>Print a String to the [Out] Parameter of the Script component.</summary>
  /// <param name="text">String to print.</param>
  private void Print(string text) { /* Implementation hidden. */ }
  /// <summary>Print a formatted String to the [Out] Parameter of the Script component.</summary>
  /// <param name="format">String format.</param>
  /// <param name="args">Formatting parameters.</param>
  private void Print(string format, params object[] args) { /* Implementation hidden. */ }
  /// <summary>Print useful information about an object instance to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj) { /* Implementation hidden. */ }
  /// <summary>Print the signatures of all the overloads of a specific method to the [Out] Parameter of the Script component. </summary>
  /// <param name="obj">Object instance to parse.</param>
  private void Reflect(object obj, string method_name) { /* Implementation hidden. */ }
#endregion

#region Members
  /// <summary>Gets the current Rhino document.</summary>
  private readonly RhinoDoc RhinoDocument;
  /// <summary>Gets the Grasshopper document that owns this script.</summary>
  private readonly GH_Document GrasshopperDocument;
  /// <summary>Gets the Grasshopper script component that owns this script.</summary>
  private readonly IGH_Component Component;
  /// <summary>
  /// Gets the current iteration count. The first call to RunScript() is associated with Iteration==0.
  /// Any subsequent call within the same solution will increment the Iteration count.
  /// </summary>
  private readonly int Iteration;
#endregion

  /// <summary>
  /// This procedure contains the user code. Input parameters are provided as regular arguments,
  /// Output parameters as ref arguments. You don't have to assign output parameters,
  /// they will have a default value.
  /// </summary>
  private void RunScript(int GMT, int Hour, int DayOfMonth, int Month, int Year, double Latitude, double Longitude, Vector3d North, Vector3d Up, ref object azimuth, ref object altitude, ref object direction)
  {
    // Implementation in C# by Mikael Nillsson
    // http://guideving.blogspot.nl/2010/08/sun-position-in-c.html
    // Adaptation to Grasshopper By Arend van Waart <Arendvw@gmail.com>
    // @see http://stackoverflow.com/questions/8708048/position-of-the-sun-given-time-of-day-latitude-and-longitude

    North.Unitize();
    Up.Unitize();


    DateTime time = new DateTime(Year, Month, DayOfMonth, Hour, 0, 0, DateTimeKind.Utc);
    time = time.AddHours(-GMT);
    this.CalculateSunPosition(time, Latitude, Longitude);
    azimuth = this.azimuth;
    altitude = this.altitude;

    // plane: x = north; y = up
    Point3d p0 = new Point3d(0, 0, 0);
    Point3d p1 = new Point3d(1, 0, 0);
    p1.Transform(Transform.Translation(North));
    p1.Transform(Transform.Rotation(this.azimuth, Up, p0));

    Plane azPlane = new Plane(p0, new Vector3d(p1), Up);
    p1.Transform(Transform.Rotation(this.altitude, azPlane.ZAxis, p0));
    Vector3d dir = new Vector3d(p1);
    dir.Unitize();
    direction = dir;
  }

  // <Custom additional code> 
  private const double Deg2Rad = Math.PI / 180.0;
  private const double Rad2Deg = 180.0 / Math.PI;
  private double altitude;
  private double azimuth;

    /*!
     * \brief Calculates the sun light.
     *
     * CalcSunPosition calculates the suns "position" based on a
     * given date and time in local time, latitude and longitude
     * expressed in decimal degrees. It is based on the method
     * found here:
     * http://www.astro.uio.no/~bgranslo/aares/calculate.html
     * The calculation is only satisfiably correct for dates in
     * the range March 1 1900 to February 28 2100.
     * \param dateTime Time and date in local time.
     * \param latitude Latitude expressed in decimal degrees.
     * \param longitude Longitude expressed in decimal degrees.
     */
  public void CalculateSunPosition(
    DateTime dateTime, double latitude, double longitude)
  {
    // Convert to UTC
    dateTime = dateTime.ToUniversalTime();

    // Number of days from J2000.0.
    double julianDate = 367 * dateTime.Year -
      (int) ((7.0 / 4.0) * (dateTime.Year +
      (int) ((dateTime.Month + 9.0) / 12.0))) +
      (int) ((275.0 * dateTime.Month) / 9.0) +
      dateTime.Day - 730531.5;

    double julianCenturies = julianDate / 36525.0;

    // Sidereal Time
    double siderealTimeHours = 6.6974 + 2400.0513 * julianCenturies;

    double siderealTimeUT = siderealTimeHours +
      (366.2422 / 365.2422) * (double) dateTime.TimeOfDay.TotalHours;

    double siderealTime = siderealTimeUT * 15 + longitude;

    // Refine to number of days (fractional) to specific time.
    julianDate += (double) dateTime.TimeOfDay.TotalHours / 24.0;
    julianCenturies = julianDate / 36525.0;

    // Solar Coordinates
    double meanLongitude = CorrectAngle(Deg2Rad *
      (280.466 + 36000.77 * julianCenturies));

    double meanAnomaly = CorrectAngle(Deg2Rad *
      (357.529 + 35999.05 * julianCenturies));

    double equationOfCenter = Deg2Rad * ((1.915 - 0.005 * julianCenturies) *
      Math.Sin(meanAnomaly) + 0.02 * Math.Sin(2 * meanAnomaly));

    double elipticalLongitude =
      CorrectAngle(meanLongitude + equationOfCenter);

    double obliquity = (23.439 - 0.013 * julianCenturies) * Deg2Rad;

    // Right Ascension
    double rightAscension = Math.Atan2(
      Math.Cos(obliquity) * Math.Sin(elipticalLongitude),
      Math.Cos(elipticalLongitude));

    double declination = Math.Asin(
      Math.Sin(rightAscension) * Math.Sin(obliquity));

    // Horizontal Coordinates
    double hourAngle = CorrectAngle(siderealTime * Deg2Rad) - rightAscension;

    if (hourAngle > Math.PI)
    {
      hourAngle -= 2 * Math.PI;
    }

    this.altitude = Math.Asin(Math.Sin(latitude * Deg2Rad) *
      Math.Sin(declination) + Math.Cos(latitude * Deg2Rad) *
      Math.Cos(declination) * Math.Cos(hourAngle));

    // Nominator and denominator for calculating Azimuth
    // angle. Needed to test which quadrant the angle is in.
    double aziNom = -Math.Sin(hourAngle);
    double aziDenom =
      Math.Tan(declination) * Math.Cos(latitude * Deg2Rad) -
      Math.Sin(latitude * Deg2Rad) * Math.Cos(hourAngle);

    this.azimuth = Math.Atan(aziNom / aziDenom);

    if (aziDenom < 0) // In 2nd or 3rd quadrant
    {
      this.azimuth += Math.PI;
    }
    else if (aziNom < 0) // In 4th quadrant
    {
      this.azimuth += 2 * Math.PI;
    }

    // Altitude
    //Console.WriteLine("Altitude: " + altitude * Rad2Deg);

    // Azimut
    //Console.WriteLine("Azimuth: " + azimuth * Rad2Deg);
  }

    /*!
    * \brief Corrects an angle.
    *
    * \param angleInRadians An angle expressed in radians.
    * \return An angle in the range 0 to 2*PI.
    */
  private static double CorrectAngle(double angleInRadians)
  {
    if (angleInRadians < 0)
    {
      return 2 * Math.PI - (Math.Abs(angleInRadians) % (2 * Math.PI));
    }
    else if (angleInRadians > 2 * Math.PI)
    {
      return angleInRadians % (2 * Math.PI);
    }
    else
    {
      return angleInRadians;
    }
  }
  // </Custom additional code> 
}