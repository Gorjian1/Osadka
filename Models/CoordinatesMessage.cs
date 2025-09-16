using System;
using System.Collections.Generic;
using System.Windows;

namespace Osadka.Messages
{
    public record CoordinatesMessage(IEnumerable<Point> Points);
}
