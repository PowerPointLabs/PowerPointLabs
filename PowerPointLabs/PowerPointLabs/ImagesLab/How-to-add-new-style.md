## How to add new style

0. Add any new properties to Model/StyleOptions.cs. For example, the new style may contain a shape, so this shape's color and transparency can be added to style option in order to control its properties.
1. Add the design effect (a basic design element) to Service/EffectsDesigner.cs. An effect is like a rectangle shape, a triangle shape, text format, textbox property etc. If it already has the effect you want, ignore this step.
2. Edit method ApplyStyle(...) of Service/StylesDesigner.cs. Tell it how to create your style by the given style option and effects.
3. Edit ModelFactory/StyleOptionsFactory.cs to generate your style's default style option (for the preview stage) and 8 default style options (for the variation stage).
4. Edit ModelFactory/StyleVariantsFactory.cs to generate your style's variants (for the variation stage).
5. Add any new string constants to TextCollection.cs, e.g. new style name, and new variant (category) name.

### Notes
0. Continuity: the new style in the preview stage should match one variation in its beginning variation stage. E.g. the new style has a default effect with color `#FFFFFF` and transparency `35` in the preview stage, so when flyout opens (beginning variation stage) it should have one variation that matches this color and transparency. This can be achieved by adjusting StyleOptionsFactory.cs and StyleVariantsFactory.cs.
1. Step-by-step Customization: all variations in the beginning variation stage should match one variant in each variant category. E.g. a variation with font color `#FFFFFF`, font size `+0`, font `Calibri`, and text position `left` will match variant `#FFFFFF` in the font color category, variant `+0` in the font size category, variant `calibri` in the font category, and variant `left` in the text position category.
