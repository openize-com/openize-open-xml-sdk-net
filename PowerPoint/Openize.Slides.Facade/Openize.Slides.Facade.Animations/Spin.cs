using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;


namespace Openize.Slides.Facade.Animations
{
    internal class Spin : IAnimation
    {
        public Timing Generate(string shapeId, int duration)
        {
            Timing timing = GenerateSpinTiming(shapeId, duration);
            return timing;
        }
        public Timing GenerateSpinTiming(String shapeId, int duration)
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

            CommonTimeNode commonTimeNode5 = new CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 8, PresetClass = TimeNodePresetClassValues.Emphasis, PresetSubtype = 0, Fill = TimeNodeFillValues.Hold, GroupId = (UInt32Value)0U, NodeType = TimeNodeValues.ClickEffect };

            StartConditionList startConditionList3 = new StartConditionList();
            Condition condition3 = new Condition() { Delay = "0" };

            startConditionList3.Append(condition3);

            ChildTimeNodeList childTimeNodeList5 = new ChildTimeNodeList();

            AnimateRotation animateRotation1 = new AnimateRotation() { By = 21600000 };

            CommonBehavior commonBehavior1 = new CommonBehavior();
            CommonTimeNode commonTimeNode6 = new CommonTimeNode() { Id = (UInt32Value)6U, Duration = duration.ToString(), Fill = TimeNodeFillValues.Hold };

            TargetElement targetElement1 = new TargetElement();
            ShapeTarget shapeTarget1 = new ShapeTarget() { ShapeId = shapeId };

            targetElement1.Append(shapeTarget1);

            AttributeNameList attributeNameList1 = new AttributeNameList();
            AttributeName attributeName1 = new AttributeName();
            attributeName1.Text = "r";

            attributeNameList1.Append(attributeName1);

            commonBehavior1.Append(commonTimeNode6);
            commonBehavior1.Append(targetElement1);
            commonBehavior1.Append(attributeNameList1);

            animateRotation1.Append(commonBehavior1);

            childTimeNodeList5.Append(animateRotation1);

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

            Condition condition4 = new Condition() { Event = TriggerEventValues.OnPrevious, Delay = "0" };

            TargetElement targetElement2 = new TargetElement();
            SlideTarget slideTarget1 = new SlideTarget();

            targetElement2.Append(slideTarget1);

            condition4.Append(targetElement2);

            previousConditionList1.Append(condition4);

            NextConditionList nextConditionList1 = new NextConditionList();

            Condition condition5 = new Condition() { Event = TriggerEventValues.OnNext, Delay = "0" };

            TargetElement targetElement3 = new TargetElement();
            SlideTarget slideTarget2 = new SlideTarget();

            targetElement3.Append(slideTarget2);

            condition5.Append(targetElement3);

            nextConditionList1.Append(condition5);

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
    }
}
