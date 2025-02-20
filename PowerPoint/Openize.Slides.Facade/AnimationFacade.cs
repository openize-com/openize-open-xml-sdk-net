using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using NonVisualGroupShapeProperties = DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Openize.Slides.Common.Enumerations;
using Openize.Slides.Common;
namespace Openize.Slides.Facade
{

    public class AnimateFacade
    {
        // Properties for ShapeId, Delay, and Duration
        public string ShapeId { get; set; }
        public int Duration { get; set; }
        public string Type { get; set; }

        // Constructor to initialize the properties
        public AnimateFacade(string shapeId = "1", int duration = 0, string type = "fade")
        {
            ShapeId = shapeId;
            Type = type;
            Duration = duration;
        }
        public Timing animate()
        {
            Timing timing = null;
            switch (Type)
            {
                case "Fade":
                    timing = GenerateFadeTiming(ShapeId, 1000);
                    break;
                case "FlyIn":
                    timing = GenerateFlyInTiming(ShapeId, 1000);
                    break;
                default:
                    timing = new Timing();
                    break;
            }
            return timing;
        }
        public Timing GenerateFadeTiming(String shapeId, int duration)
        {
            Timing timing1 = new Timing();

            TimeNodeList timeNodeList1 = new TimeNodeList();

            ParallelTimeNode parallelTimeNode1 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode1 = new CommonTimeNode() { Id = (UInt32Value)1U, Duration = "indefinite", Restart = TimeNodeRestartValues.Never, NodeType = TimeNodeValues.TmingRoot };

            ChildTimeNodeList childTimeNodeList1 = new ChildTimeNodeList();

            SequenceTimeNode sequenceTimeNode1 = new SequenceTimeNode() { Concurrent = true, NextAction = NextActionValues.Seek };

            CommonTimeNode commonTimeNode2 = new CommonTimeNode() { Id = (UInt32Value)2U, Duration = "indefinite", NodeType = TimeNodeValues.MainSequence };

            ChildTimeNodeList childTimeNodeList2 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode2 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode3 = new CommonTimeNode() { Id = (UInt32Value)3U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList1 = new StartConditionList();
            Condition condition1 = new Condition() { Delay = "indefinite" };

            startConditionList1.Append(condition1);

            ChildTimeNodeList childTimeNodeList3 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode3 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode4 = new CommonTimeNode() { Id = (UInt32Value)4U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList2 = new StartConditionList();
            Condition condition2 = new Condition() { Delay = "0" };

            startConditionList2.Append(condition2);

            ChildTimeNodeList childTimeNodeList4 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode4 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode5 = new CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 42, PresetClass = TimeNodePresetClassValues.Entrance, PresetSubtype = 0, Fill = TimeNodeFillValues.Hold, GroupId = (UInt32Value)0U, NodeType = TimeNodeValues.ClickEffect };

            StartConditionList startConditionList3 = new StartConditionList();
            Condition condition3 = new Condition() { Delay = "0" };

            startConditionList3.Append(condition3);

            ChildTimeNodeList childTimeNodeList5 = new ChildTimeNodeList();

            SetBehavior setBehavior1 = new SetBehavior();

            CommonBehavior commonBehavior1 = new CommonBehavior();

            CommonTimeNode commonTimeNode6 = new CommonTimeNode() { Id = (UInt32Value)6U, Duration = "1", Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList4 = new StartConditionList();
            Condition condition4 = new Condition() { Delay = "0" };

            startConditionList4.Append(condition4);

            commonTimeNode6.Append(startConditionList4);

            TargetElement targetElement1 = new TargetElement();
            ShapeTarget shapeTarget1 = new ShapeTarget() { ShapeId = shapeId };

            targetElement1.Append(shapeTarget1);

            AttributeNameList attributeNameList1 = new AttributeNameList();
            AttributeName attributeName1 = new AttributeName();
            attributeName1.Text = "style.visibility";

            attributeNameList1.Append(attributeName1);

            commonBehavior1.Append(commonTimeNode6);
            commonBehavior1.Append(targetElement1);
            commonBehavior1.Append(attributeNameList1);

            ToVariantValue toVariantValue1 = new ToVariantValue();
            StringVariantValue stringVariantValue1 = new StringVariantValue() { Val = "visible" };

            toVariantValue1.Append(stringVariantValue1);

            setBehavior1.Append(commonBehavior1);
            setBehavior1.Append(toVariantValue1);

            AnimateEffect animateEffect1 = new AnimateEffect() { Transition = AnimateEffectTransitionValues.In, Filter = "fade" };

            CommonBehavior commonBehavior2 = new CommonBehavior();
            CommonTimeNode commonTimeNode7 = new CommonTimeNode() { Id = (UInt32Value)7U, Duration = duration.ToString() };

            TargetElement targetElement2 = new TargetElement();
            ShapeTarget shapeTarget2 = new ShapeTarget() { ShapeId = shapeId };

            targetElement2.Append(shapeTarget2);

            commonBehavior2.Append(commonTimeNode7);
            commonBehavior2.Append(targetElement2);

            animateEffect1.Append(commonBehavior2);

            Animate animate1 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

            CommonBehavior commonBehavior3 = new CommonBehavior();
            CommonTimeNode commonTimeNode8 = new CommonTimeNode() { Id = (UInt32Value)8U, Duration = duration.ToString(), Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement3 = new TargetElement();
            ShapeTarget shapeTarget3 = new ShapeTarget() { ShapeId = shapeId };

            targetElement3.Append(shapeTarget3);

            AttributeNameList attributeNameList2 = new AttributeNameList();
            AttributeName attributeName2 = new AttributeName();
            attributeName2.Text = "ppt_x";

            attributeNameList2.Append(attributeName2);

            commonBehavior3.Append(commonTimeNode8);
            commonBehavior3.Append(targetElement3);
            commonBehavior3.Append(attributeNameList2);

            TimeAnimateValueList timeAnimateValueList1 = new TimeAnimateValueList();

            TimeAnimateValue timeAnimateValue1 = new TimeAnimateValue() { Time = "0" };

            VariantValue variantValue1 = new VariantValue();
            StringVariantValue stringVariantValue2 = new StringVariantValue() { Val = "#ppt_x" };

            variantValue1.Append(stringVariantValue2);

            timeAnimateValue1.Append(variantValue1);

            TimeAnimateValue timeAnimateValue2 = new TimeAnimateValue() { Time = "100000" };

            VariantValue variantValue2 = new VariantValue();
            StringVariantValue stringVariantValue3 = new StringVariantValue() { Val = "#ppt_x" };

            variantValue2.Append(stringVariantValue3);

            timeAnimateValue2.Append(variantValue2);

            timeAnimateValueList1.Append(timeAnimateValue1);
            timeAnimateValueList1.Append(timeAnimateValue2);

            animate1.Append(commonBehavior3);
            animate1.Append(timeAnimateValueList1);

            Animate animate2 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

            CommonBehavior commonBehavior4 = new CommonBehavior();
            CommonTimeNode commonTimeNode9 = new CommonTimeNode() { Id = (UInt32Value)9U, Duration = duration.ToString(), Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement4 = new TargetElement();
            ShapeTarget shapeTarget4 = new ShapeTarget() { ShapeId = shapeId };

            targetElement4.Append(shapeTarget4);

            AttributeNameList attributeNameList3 = new AttributeNameList();
            AttributeName attributeName3 = new AttributeName();
            attributeName3.Text = "ppt_y";

            attributeNameList3.Append(attributeName3);

            commonBehavior4.Append(commonTimeNode9);
            commonBehavior4.Append(targetElement4);
            commonBehavior4.Append(attributeNameList3);

            TimeAnimateValueList timeAnimateValueList2 = new TimeAnimateValueList();

            TimeAnimateValue timeAnimateValue3 = new TimeAnimateValue() { Time = "0" };

            VariantValue variantValue3 = new VariantValue();
            StringVariantValue stringVariantValue4 = new StringVariantValue() { Val = "#ppt_y+.1" };

            variantValue3.Append(stringVariantValue4);

            timeAnimateValue3.Append(variantValue3);

            TimeAnimateValue timeAnimateValue4 = new TimeAnimateValue() { Time = "100000" };

            VariantValue variantValue4 = new VariantValue();
            StringVariantValue stringVariantValue5 = new StringVariantValue() { Val = "#ppt_y" };

            variantValue4.Append(stringVariantValue5);

            timeAnimateValue4.Append(variantValue4);

            timeAnimateValueList2.Append(timeAnimateValue3);
            timeAnimateValueList2.Append(timeAnimateValue4);

            animate2.Append(commonBehavior4);
            animate2.Append(timeAnimateValueList2);

            childTimeNodeList5.Append(setBehavior1);
            childTimeNodeList5.Append(animateEffect1);
            childTimeNodeList5.Append(animate1);
            childTimeNodeList5.Append(animate2);

            commonTimeNode5.Append(startConditionList3);
            commonTimeNode5.Append(childTimeNodeList5);

            parallelTimeNode4.Append(commonTimeNode5);

            childTimeNodeList4.Append(parallelTimeNode4);

            commonTimeNode4.Append(startConditionList2);
            commonTimeNode4.Append(childTimeNodeList4);

            parallelTimeNode3.Append(commonTimeNode4);

            childTimeNodeList3.Append(parallelTimeNode3);

            commonTimeNode3.Append(startConditionList1);
            commonTimeNode3.Append(childTimeNodeList3);

            parallelTimeNode2.Append(commonTimeNode3);

            childTimeNodeList2.Append(parallelTimeNode2);

            commonTimeNode2.Append(childTimeNodeList2);

            PreviousConditionList previousConditionList1 = new PreviousConditionList();

            Condition condition5 = new Condition() { Event = TriggerEventValues.OnPrevious, Delay = "0" };

            TargetElement targetElement5 = new TargetElement();
            SlideTarget slideTarget1 = new SlideTarget();

            targetElement5.Append(slideTarget1);

            condition5.Append(targetElement5);

            previousConditionList1.Append(condition5);

            NextConditionList nextConditionList1 = new NextConditionList();

            Condition condition6 = new Condition() { Event = TriggerEventValues.OnNext, Delay = "0" };

            TargetElement targetElement6 = new TargetElement();
            SlideTarget slideTarget2 = new SlideTarget();

            targetElement6.Append(slideTarget2);

            condition6.Append(targetElement6);

            nextConditionList1.Append(condition6);

            sequenceTimeNode1.Append(commonTimeNode2);
            sequenceTimeNode1.Append(previousConditionList1);
            sequenceTimeNode1.Append(nextConditionList1);

            childTimeNodeList1.Append(sequenceTimeNode1);

            commonTimeNode1.Append(childTimeNodeList1);

            parallelTimeNode1.Append(commonTimeNode1);

            timeNodeList1.Append(parallelTimeNode1);

            BuildList buildList1 = new BuildList();
            BuildParagraph buildParagraph1 = new BuildParagraph() { ShapeId = shapeId, GroupId = (UInt32Value)0U, AnimateBackground = true };

            buildList1.Append(buildParagraph1);

            timing1.Append(timeNodeList1);
            timing1.Append(buildList1);
            return timing1;
        }

        public Timing GenerateFlyInTiming(String shapeId, int duration)
        {
            Timing timing1 = new Timing();

            TimeNodeList timeNodeList1 = new TimeNodeList();

            ParallelTimeNode parallelTimeNode1 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode1 = new CommonTimeNode() { Id = (UInt32Value)1U, Duration = "indefinite", Restart = TimeNodeRestartValues.Never, NodeType = TimeNodeValues.TmingRoot };

            ChildTimeNodeList childTimeNodeList1 = new ChildTimeNodeList();

            SequenceTimeNode sequenceTimeNode1 = new SequenceTimeNode() { Concurrent = true, NextAction = NextActionValues.Seek };

            CommonTimeNode commonTimeNode2 = new CommonTimeNode() { Id = (UInt32Value)2U, Duration = "indefinite", NodeType = TimeNodeValues.MainSequence };

            ChildTimeNodeList childTimeNodeList2 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode2 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode3 = new CommonTimeNode() { Id = (UInt32Value)3U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList1 = new StartConditionList();
            Condition condition1 = new Condition() { Delay = "indefinite" };

            startConditionList1.Append(condition1);

            ChildTimeNodeList childTimeNodeList3 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode3 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode4 = new CommonTimeNode() { Id = (UInt32Value)4U, Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList2 = new StartConditionList();
            Condition condition2 = new Condition() { Delay = "0" };

            startConditionList2.Append(condition2);

            ChildTimeNodeList childTimeNodeList4 = new ChildTimeNodeList();

            ParallelTimeNode parallelTimeNode4 = new ParallelTimeNode();

            CommonTimeNode commonTimeNode5 = new CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 2, PresetClass = TimeNodePresetClassValues.Entrance, PresetSubtype = 4, Fill = TimeNodeFillValues.Hold, GroupId = (UInt32Value)0U, NodeType = TimeNodeValues.ClickEffect };

            StartConditionList startConditionList3 = new StartConditionList();
            Condition condition3 = new Condition() { Delay = "0" };

            startConditionList3.Append(condition3);

            ChildTimeNodeList childTimeNodeList5 = new ChildTimeNodeList();

            SetBehavior setBehavior1 = new SetBehavior();

            CommonBehavior commonBehavior1 = new CommonBehavior();

            CommonTimeNode commonTimeNode6 = new CommonTimeNode() { Id = (UInt32Value)6U, Duration = "1", Fill = TimeNodeFillValues.Hold };

            StartConditionList startConditionList4 = new StartConditionList();
            Condition condition4 = new Condition() { Delay = "0" };

            startConditionList4.Append(condition4);

            commonTimeNode6.Append(startConditionList4);

            TargetElement targetElement1 = new TargetElement();
            ShapeTarget shapeTarget1 = new ShapeTarget() { ShapeId = ShapeId };

            targetElement1.Append(shapeTarget1);

            AttributeNameList attributeNameList1 = new AttributeNameList();
            AttributeName attributeName1 = new AttributeName();
            attributeName1.Text = "style.visibility";

            attributeNameList1.Append(attributeName1);

            commonBehavior1.Append(commonTimeNode6);
            commonBehavior1.Append(targetElement1);
            commonBehavior1.Append(attributeNameList1);

            ToVariantValue toVariantValue1 = new ToVariantValue();
            StringVariantValue stringVariantValue1 = new StringVariantValue() { Val = "visible" };

            toVariantValue1.Append(stringVariantValue1);

            setBehavior1.Append(commonBehavior1);
            setBehavior1.Append(toVariantValue1);

            Animate animate1 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

            CommonBehavior commonBehavior2 = new CommonBehavior() { Additive = BehaviorAdditiveValues.Base };
            CommonTimeNode commonTimeNode7 = new CommonTimeNode() { Id = (UInt32Value)7U, Duration = duration.ToString(), Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement2 = new TargetElement();
            ShapeTarget shapeTarget2 = new ShapeTarget() { ShapeId = ShapeId };

            targetElement2.Append(shapeTarget2);

            AttributeNameList attributeNameList2 = new AttributeNameList();
            AttributeName attributeName2 = new AttributeName();
            attributeName2.Text = "ppt_x";

            attributeNameList2.Append(attributeName2);

            commonBehavior2.Append(commonTimeNode7);
            commonBehavior2.Append(targetElement2);
            commonBehavior2.Append(attributeNameList2);

            TimeAnimateValueList timeAnimateValueList1 = new TimeAnimateValueList();

            TimeAnimateValue timeAnimateValue1 = new TimeAnimateValue() { Time = "0" };

            VariantValue variantValue1 = new VariantValue();
            StringVariantValue stringVariantValue2 = new StringVariantValue() { Val = "#ppt_x" };

            variantValue1.Append(stringVariantValue2);

            timeAnimateValue1.Append(variantValue1);

            TimeAnimateValue timeAnimateValue2 = new TimeAnimateValue() { Time = "100000" };

            VariantValue variantValue2 = new VariantValue();
            StringVariantValue stringVariantValue3 = new StringVariantValue() { Val = "#ppt_x" };

            variantValue2.Append(stringVariantValue3);

            timeAnimateValue2.Append(variantValue2);

            timeAnimateValueList1.Append(timeAnimateValue1);
            timeAnimateValueList1.Append(timeAnimateValue2);

            animate1.Append(commonBehavior2);
            animate1.Append(timeAnimateValueList1);

            Animate animate2 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

            CommonBehavior commonBehavior3 = new CommonBehavior() { Additive = BehaviorAdditiveValues.Base };
            CommonTimeNode commonTimeNode8 = new CommonTimeNode() { Id = (UInt32Value)8U, Duration = Duration.ToString(), Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement3 = new TargetElement();
            ShapeTarget shapeTarget3 = new ShapeTarget() { ShapeId = ShapeId };

            targetElement3.Append(shapeTarget3);

            AttributeNameList attributeNameList3 = new AttributeNameList();
            AttributeName attributeName3 = new AttributeName();
            attributeName3.Text = "ppt_y";

            attributeNameList3.Append(attributeName3);

            commonBehavior3.Append(commonTimeNode8);
            commonBehavior3.Append(targetElement3);
            commonBehavior3.Append(attributeNameList3);

            TimeAnimateValueList timeAnimateValueList2 = new TimeAnimateValueList();

            TimeAnimateValue timeAnimateValue3 = new TimeAnimateValue() { Time = "0" };

            VariantValue variantValue3 = new VariantValue();
            StringVariantValue stringVariantValue4 = new StringVariantValue() { Val = "1+#ppt_h/2" };

            variantValue3.Append(stringVariantValue4);

            timeAnimateValue3.Append(variantValue3);

            TimeAnimateValue timeAnimateValue4 = new TimeAnimateValue() { Time = "100000" };

            VariantValue variantValue4 = new VariantValue();
            StringVariantValue stringVariantValue5 = new StringVariantValue() { Val = "#ppt_y" };

            variantValue4.Append(stringVariantValue5);

            timeAnimateValue4.Append(variantValue4);

            timeAnimateValueList2.Append(timeAnimateValue3);
            timeAnimateValueList2.Append(timeAnimateValue4);

            animate2.Append(commonBehavior3);
            animate2.Append(timeAnimateValueList2);

            childTimeNodeList5.Append(setBehavior1);
            childTimeNodeList5.Append(animate1);
            childTimeNodeList5.Append(animate2);

            commonTimeNode5.Append(startConditionList3);
            commonTimeNode5.Append(childTimeNodeList5);

            parallelTimeNode4.Append(commonTimeNode5);

            childTimeNodeList4.Append(parallelTimeNode4);

            commonTimeNode4.Append(startConditionList2);
            commonTimeNode4.Append(childTimeNodeList4);

            parallelTimeNode3.Append(commonTimeNode4);

            childTimeNodeList3.Append(parallelTimeNode3);

            commonTimeNode3.Append(startConditionList1);
            commonTimeNode3.Append(childTimeNodeList3);

            parallelTimeNode2.Append(commonTimeNode3);

            childTimeNodeList2.Append(parallelTimeNode2);

            commonTimeNode2.Append(childTimeNodeList2);

            PreviousConditionList previousConditionList1 = new PreviousConditionList();

            Condition condition5 = new Condition() { Event = TriggerEventValues.OnPrevious, Delay = "0" };

            TargetElement targetElement4 = new TargetElement();
            SlideTarget slideTarget1 = new SlideTarget();

            targetElement4.Append(slideTarget1);

            condition5.Append(targetElement4);

            previousConditionList1.Append(condition5);

            NextConditionList nextConditionList1 = new NextConditionList();

            Condition condition6 = new Condition() { Event = TriggerEventValues.OnNext, Delay = "0" };

            TargetElement targetElement5 = new TargetElement();
            SlideTarget slideTarget2 = new SlideTarget();

            targetElement5.Append(slideTarget2);

            condition6.Append(targetElement5);

            nextConditionList1.Append(condition6);

            sequenceTimeNode1.Append(commonTimeNode2);
            sequenceTimeNode1.Append(previousConditionList1);
            sequenceTimeNode1.Append(nextConditionList1);

            childTimeNodeList1.Append(sequenceTimeNode1);

            commonTimeNode1.Append(childTimeNodeList1);

            parallelTimeNode1.Append(commonTimeNode1);

            timeNodeList1.Append(parallelTimeNode1);

            BuildList buildList1 = new BuildList();
            BuildParagraph buildParagraph1 = new BuildParagraph() { ShapeId = ShapeId, GroupId = (UInt32Value)0U, AnimateBackground = true };

            buildList1.Append(buildParagraph1);

            timing1.Append(timeNodeList1);
            timing1.Append(buildList1);
            return timing1;
        }

    }


}
