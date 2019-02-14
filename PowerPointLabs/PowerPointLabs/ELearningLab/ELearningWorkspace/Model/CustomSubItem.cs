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

        private Shape shape;
        private Effect effect;
        private string shapeId;
        private string shapeName;
        private AnimationType type;

        public CustomSubItem(Shape shape, Effect effect)
        {
            this.shape = shape;
            this.effect = effect;
            shapeId = shape.Id.ToString();
            shapeName = shape.Name;
            type = EffectToAnimationTypeConverter.GetAnimationTypeOfEffect(effect);
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
            return shape.Equals(other.shape) && effect.Equals(other.effect);
        }

        public override int GetHashCode()
        {
            var hashCode = -1136360337;
            hashCode = hashCode * -1521134295 + Type.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<Shape>.Default.GetHashCode(shape);
            hashCode = hashCode * -1521134295 + EqualityComparer<Effect>.Default.GetHashCode(effect);
            return hashCode;
        }
    }
}
