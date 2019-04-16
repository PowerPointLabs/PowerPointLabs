using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Converters;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class CustomSubItem: IEquatable<CustomSubItem>
    {
        public string ShapeName
        {
            get
            {
                return shapeName;
            }
        }

        public string ShapeId
        {
            get
            {
                return shapeId;
            }
        }

        public AnimationType Type
        {
            get
            {
                return type;
            }
        }

        private string shapeId;
        private string shapeName;
        private AnimationType type;

        public CustomSubItem(string shapeName, string shapeId, AnimationType type)
        {
            this.shapeId = shapeId;
            this.shapeName = shapeName;
            this.type = type;
        }

        public override bool Equals(object other)
        {
            if (other == null || other.GetType() != GetType())
            {
                return false;
            }

            if (ReferenceEquals(other, this))
            {
                return true;
            }
            return Equals(other as CustomSubItem);
        }

        public bool Equals(CustomSubItem other)
        {
            return shapeName.Equals(other.shapeName) && shapeId.Equals(other.shapeId)
                && type.Equals(other.type);
        }

        public override int GetHashCode()
        {
            var hashCode = -1136360337;
            hashCode = hashCode * -1521134295 + Type.GetHashCode();
            hashCode = hashCode * -1521134295 + shapeName.GetHashCode();
            hashCode = hashCode * -1521134295 + shapeId.GetHashCode();
            return hashCode;
        }
    }
}
