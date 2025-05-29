using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;


namespace Openize.Slides.Facade.Animations
{
    internal class Bounce : IAnimation
    {
        public Timing Generate(string shapeId, int duration)
        {
            Timing timing = GenerateBounceTiming(shapeId, duration);
           return timing;
        }
           public Timing GenerateBounceTiming(String shapeId, int duration)
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

                CommonTimeNode commonTimeNode5 = new CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 26, PresetClass = TimeNodePresetClassValues.Entrance, PresetSubtype = 0, Fill = TimeNodeFillValues.Hold, GroupId = (UInt32Value)0U, NodeType = TimeNodeValues.ClickEffect };

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

                AnimateEffect animateEffect1 = new AnimateEffect() { Transition = AnimateEffectTransitionValues.In, Filter = "wipe(down)" };

                CommonBehavior commonBehavior2 = new CommonBehavior();

                CommonTimeNode commonTimeNode7 = new CommonTimeNode() { Id = (UInt32Value)7U, Duration = "580" };

                StartConditionList startConditionList5 = new StartConditionList();
                Condition condition5 = new Condition() { Delay = "0" };

                startConditionList5.Append(condition5);

                commonTimeNode7.Append(startConditionList5);

                TargetElement targetElement2 = new TargetElement();
                ShapeTarget shapeTarget2 = new ShapeTarget() { ShapeId = shapeId };

                targetElement2.Append(shapeTarget2);

                commonBehavior2.Append(commonTimeNode7);
                commonBehavior2.Append(targetElement2);

                animateEffect1.Append(commonBehavior2);

                Animate animate1 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

                CommonBehavior commonBehavior3 = new CommonBehavior();

                CommonTimeNode commonTimeNode8 = new CommonTimeNode() { Id = (UInt32Value)8U, Duration = "1822", TimeFilter = "0,0; 0.14,0.36; 0.43,0.73; 0.71,0.91; 1.0,1.0" };

                StartConditionList startConditionList6 = new StartConditionList();
                Condition condition6 = new Condition() { Delay = "0" };

                startConditionList6.Append(condition6);

                commonTimeNode8.Append(startConditionList6);

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
                StringVariantValue stringVariantValue2 = new StringVariantValue() { Val = "#ppt_x-0.25" };

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

                CommonTimeNode commonTimeNode9 = new CommonTimeNode() { Id = (UInt32Value)9U, Duration = "664", TimeFilter = "0.0,0.0; 0.25,0.07; 0.50,0.2; 0.75,0.467; 1.0,1.0" };

                StartConditionList startConditionList7 = new StartConditionList();
                Condition condition7 = new Condition() { Delay = "0" };

                startConditionList7.Append(condition7);

                commonTimeNode9.Append(startConditionList7);

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

                TimeAnimateValue timeAnimateValue3 = new TimeAnimateValue() { Time = "0", Fomula = "#ppt_y-sin(pi*$)/3" };

                VariantValue variantValue3 = new VariantValue();
                FloatVariantValue floatVariantValue1 = new FloatVariantValue() { Val = 0.5F };

                variantValue3.Append(floatVariantValue1);

                timeAnimateValue3.Append(variantValue3);

                TimeAnimateValue timeAnimateValue4 = new TimeAnimateValue() { Time = "100000" };

                VariantValue variantValue4 = new VariantValue();
                FloatVariantValue floatVariantValue2 = new FloatVariantValue() { Val = 1F };

                variantValue4.Append(floatVariantValue2);

                timeAnimateValue4.Append(variantValue4);

                timeAnimateValueList2.Append(timeAnimateValue3);
                timeAnimateValueList2.Append(timeAnimateValue4);

                animate2.Append(commonBehavior4);
                animate2.Append(timeAnimateValueList2);

                Animate animate3 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

                CommonBehavior commonBehavior5 = new CommonBehavior();

                CommonTimeNode commonTimeNode10 = new CommonTimeNode() { Id = (UInt32Value)10U, Duration = "664", TimeFilter = "0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1" };

                StartConditionList startConditionList8 = new StartConditionList();
                Condition condition8 = new Condition() { Delay = "664" };

                startConditionList8.Append(condition8);

                commonTimeNode10.Append(startConditionList8);

                TargetElement targetElement5 = new TargetElement();
                ShapeTarget shapeTarget5 = new ShapeTarget() { ShapeId = shapeId };

                targetElement5.Append(shapeTarget5);

                AttributeNameList attributeNameList4 = new AttributeNameList();
                AttributeName attributeName4 = new AttributeName();
                attributeName4.Text = "ppt_y";

                attributeNameList4.Append(attributeName4);

                commonBehavior5.Append(commonTimeNode10);
                commonBehavior5.Append(targetElement5);
                commonBehavior5.Append(attributeNameList4);

                TimeAnimateValueList timeAnimateValueList3 = new TimeAnimateValueList();

                TimeAnimateValue timeAnimateValue5 = new TimeAnimateValue() { Time = "0", Fomula = "#ppt_y-sin(pi*$)/9" };

                VariantValue variantValue5 = new VariantValue();
                FloatVariantValue floatVariantValue3 = new FloatVariantValue() { Val = 0F };

                variantValue5.Append(floatVariantValue3);

                timeAnimateValue5.Append(variantValue5);

                TimeAnimateValue timeAnimateValue6 = new TimeAnimateValue() { Time = "100000" };

                VariantValue variantValue6 = new VariantValue();
                FloatVariantValue floatVariantValue4 = new FloatVariantValue() { Val = 1F };

                variantValue6.Append(floatVariantValue4);

                timeAnimateValue6.Append(variantValue6);

                timeAnimateValueList3.Append(timeAnimateValue5);
                timeAnimateValueList3.Append(timeAnimateValue6);

                animate3.Append(commonBehavior5);
                animate3.Append(timeAnimateValueList3);

                Animate animate4 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

                CommonBehavior commonBehavior6 = new CommonBehavior();

                CommonTimeNode commonTimeNode11 = new CommonTimeNode() { Id = (UInt32Value)11U, Duration = "332", TimeFilter = "0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1" };

                StartConditionList startConditionList9 = new StartConditionList();
                Condition condition9 = new Condition() { Delay = "1324" };

                startConditionList9.Append(condition9);

                commonTimeNode11.Append(startConditionList9);

                TargetElement targetElement6 = new TargetElement();
                ShapeTarget shapeTarget6 = new ShapeTarget() { ShapeId = shapeId };

                targetElement6.Append(shapeTarget6);

                AttributeNameList attributeNameList5 = new AttributeNameList();
                AttributeName attributeName5 = new AttributeName();
                attributeName5.Text = "ppt_y";

                attributeNameList5.Append(attributeName5);

                commonBehavior6.Append(commonTimeNode11);
                commonBehavior6.Append(targetElement6);
                commonBehavior6.Append(attributeNameList5);

                TimeAnimateValueList timeAnimateValueList4 = new TimeAnimateValueList();

                TimeAnimateValue timeAnimateValue7 = new TimeAnimateValue() { Time = "0", Fomula = "#ppt_y-sin(pi*$)/27" };

                VariantValue variantValue7 = new VariantValue();
                FloatVariantValue floatVariantValue5 = new FloatVariantValue() { Val = 0F };

                variantValue7.Append(floatVariantValue5);

                timeAnimateValue7.Append(variantValue7);

                TimeAnimateValue timeAnimateValue8 = new TimeAnimateValue() { Time = "100000" };

                VariantValue variantValue8 = new VariantValue();
                FloatVariantValue floatVariantValue6 = new FloatVariantValue() { Val = 1F };

                variantValue8.Append(floatVariantValue6);

                timeAnimateValue8.Append(variantValue8);

                timeAnimateValueList4.Append(timeAnimateValue7);
                timeAnimateValueList4.Append(timeAnimateValue8);

                animate4.Append(commonBehavior6);
                animate4.Append(timeAnimateValueList4);

                Animate animate5 = new Animate() { CalculationMode = AnimateBehaviorCalculateModeValues.Linear, ValueType = AnimateBehaviorValues.Number };

                CommonBehavior commonBehavior7 = new CommonBehavior();

                CommonTimeNode commonTimeNode12 = new CommonTimeNode() { Id = (UInt32Value)12U, Duration = "164", TimeFilter = "0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1" };

                StartConditionList startConditionList10 = new StartConditionList();
                Condition condition10 = new Condition() { Delay = "1656" };

                startConditionList10.Append(condition10);

                commonTimeNode12.Append(startConditionList10);

                TargetElement targetElement7 = new TargetElement();
                ShapeTarget shapeTarget7 = new ShapeTarget() { ShapeId = shapeId };

                targetElement7.Append(shapeTarget7);

                AttributeNameList attributeNameList6 = new AttributeNameList();
                AttributeName attributeName6 = new AttributeName();
                attributeName6.Text = "ppt_y";

                attributeNameList6.Append(attributeName6);

                commonBehavior7.Append(commonTimeNode12);
                commonBehavior7.Append(targetElement7);
                commonBehavior7.Append(attributeNameList6);

                TimeAnimateValueList timeAnimateValueList5 = new TimeAnimateValueList();

                TimeAnimateValue timeAnimateValue9 = new TimeAnimateValue() { Time = "0", Fomula = "#ppt_y-sin(pi*$)/81" };

                VariantValue variantValue9 = new VariantValue();
                FloatVariantValue floatVariantValue7 = new FloatVariantValue() { Val = 0F };

                variantValue9.Append(floatVariantValue7);

                timeAnimateValue9.Append(variantValue9);

                TimeAnimateValue timeAnimateValue10 = new TimeAnimateValue() { Time = "100000" };

                VariantValue variantValue10 = new VariantValue();
                FloatVariantValue floatVariantValue8 = new FloatVariantValue() { Val = 1F };

                variantValue10.Append(floatVariantValue8);

                timeAnimateValue10.Append(variantValue10);

                timeAnimateValueList5.Append(timeAnimateValue9);
                timeAnimateValueList5.Append(timeAnimateValue10);

                animate5.Append(commonBehavior7);
                animate5.Append(timeAnimateValueList5);

                AnimateScale animateScale1 = new AnimateScale();

                CommonBehavior commonBehavior8 = new CommonBehavior();

                CommonTimeNode commonTimeNode13 = new CommonTimeNode() { Id = (UInt32Value)13U, Duration = "26" };

                StartConditionList startConditionList11 = new StartConditionList();
                Condition condition11 = new Condition() { Delay = "650" };

                startConditionList11.Append(condition11);

                commonTimeNode13.Append(startConditionList11);

                TargetElement targetElement8 = new TargetElement();
                ShapeTarget shapeTarget8 = new ShapeTarget() { ShapeId = shapeId };

                targetElement8.Append(shapeTarget8);

                commonBehavior8.Append(commonTimeNode13);
                commonBehavior8.Append(targetElement8);
                ToPosition toPosition1 = new ToPosition() { X = 100000, Y = 60000 };

                animateScale1.Append(commonBehavior8);
                animateScale1.Append(toPosition1);

                AnimateScale animateScale2 = new AnimateScale();

                CommonBehavior commonBehavior9 = new CommonBehavior();

                CommonTimeNode commonTimeNode14 = new CommonTimeNode() { Id = (UInt32Value)14U, Duration = "166", Deceleration = 50000 };

                StartConditionList startConditionList12 = new StartConditionList();
                Condition condition12 = new Condition() { Delay = "676" };

                startConditionList12.Append(condition12);

                commonTimeNode14.Append(startConditionList12);

                TargetElement targetElement9 = new TargetElement();
                ShapeTarget shapeTarget9 = new ShapeTarget() { ShapeId = shapeId };

                targetElement9.Append(shapeTarget9);

                commonBehavior9.Append(commonTimeNode14);
                commonBehavior9.Append(targetElement9);
                ToPosition toPosition2 = new ToPosition() { X = 100000, Y = 100000 };

                animateScale2.Append(commonBehavior9);
                animateScale2.Append(toPosition2);

                AnimateScale animateScale3 = new AnimateScale();

                CommonBehavior commonBehavior10 = new CommonBehavior();

                CommonTimeNode commonTimeNode15 = new CommonTimeNode() { Id = (UInt32Value)15U, Duration = "26" };

                StartConditionList startConditionList13 = new StartConditionList();
                Condition condition13 = new Condition() { Delay = "1312" };

                startConditionList13.Append(condition13);

                commonTimeNode15.Append(startConditionList13);

                TargetElement targetElement10 = new TargetElement();
                ShapeTarget shapeTarget10 = new ShapeTarget() { ShapeId = shapeId };

                targetElement10.Append(shapeTarget10);

                commonBehavior10.Append(commonTimeNode15);
                commonBehavior10.Append(targetElement10);
                ToPosition toPosition3 = new ToPosition() { X = 100000, Y = 80000 };

                animateScale3.Append(commonBehavior10);
                animateScale3.Append(toPosition3);

                AnimateScale animateScale4 = new AnimateScale();

                CommonBehavior commonBehavior11 = new CommonBehavior();

                CommonTimeNode commonTimeNode16 = new CommonTimeNode() { Id = (UInt32Value)16U, Duration = "166", Deceleration = 50000 };

                StartConditionList startConditionList14 = new StartConditionList();
                Condition condition14 = new Condition() { Delay = "1338" };

                startConditionList14.Append(condition14);

                commonTimeNode16.Append(startConditionList14);

                TargetElement targetElement11 = new TargetElement();
                ShapeTarget shapeTarget11 = new ShapeTarget() { ShapeId = shapeId };

                targetElement11.Append(shapeTarget11);

                commonBehavior11.Append(commonTimeNode16);
                commonBehavior11.Append(targetElement11);
                ToPosition toPosition4 = new ToPosition() { X = 100000, Y = 100000 };

                animateScale4.Append(commonBehavior11);
                animateScale4.Append(toPosition4);

                AnimateScale animateScale5 = new AnimateScale();

                CommonBehavior commonBehavior12 = new CommonBehavior();

                CommonTimeNode commonTimeNode17 = new CommonTimeNode() { Id = (UInt32Value)17U, Duration = "26" };

                StartConditionList startConditionList15 = new StartConditionList();
                Condition condition15 = new Condition() { Delay = "1642" };

                startConditionList15.Append(condition15);

                commonTimeNode17.Append(startConditionList15);

                TargetElement targetElement12 = new TargetElement();
                ShapeTarget shapeTarget12 = new ShapeTarget() { ShapeId = shapeId };

                targetElement12.Append(shapeTarget12);

                commonBehavior12.Append(commonTimeNode17);
                commonBehavior12.Append(targetElement12);
                ToPosition toPosition5 = new ToPosition() { X = 100000, Y = 90000 };

                animateScale5.Append(commonBehavior12);
                animateScale5.Append(toPosition5);

                AnimateScale animateScale6 = new AnimateScale();

                CommonBehavior commonBehavior13 = new CommonBehavior();

                CommonTimeNode commonTimeNode18 = new CommonTimeNode() { Id = (UInt32Value)18U, Duration = "166", Deceleration = 50000 };

                StartConditionList startConditionList16 = new StartConditionList();
                Condition condition16 = new Condition() { Delay = "1668" };

                startConditionList16.Append(condition16);

                commonTimeNode18.Append(startConditionList16);

                TargetElement targetElement13 = new TargetElement();
                ShapeTarget shapeTarget13 = new ShapeTarget() { ShapeId = shapeId };

                targetElement13.Append(shapeTarget13);

                commonBehavior13.Append(commonTimeNode18);
                commonBehavior13.Append(targetElement13);
                ToPosition toPosition6 = new ToPosition() { X = 100000, Y = 100000 };

                animateScale6.Append(commonBehavior13);
                animateScale6.Append(toPosition6);

                AnimateScale animateScale7 = new AnimateScale();

                CommonBehavior commonBehavior14 = new CommonBehavior();

                CommonTimeNode commonTimeNode19 = new CommonTimeNode() { Id = (UInt32Value)19U, Duration = "26" };

                StartConditionList startConditionList17 = new StartConditionList();
                Condition condition17 = new Condition() { Delay = "1808" };

                startConditionList17.Append(condition17);

                commonTimeNode19.Append(startConditionList17);

                TargetElement targetElement14 = new TargetElement();
                ShapeTarget shapeTarget14 = new ShapeTarget() { ShapeId = shapeId };

                targetElement14.Append(shapeTarget14);

                commonBehavior14.Append(commonTimeNode19);
                commonBehavior14.Append(targetElement14);
                ToPosition toPosition7 = new ToPosition() { X = 100000, Y = 95000 };

                animateScale7.Append(commonBehavior14);
                animateScale7.Append(toPosition7);

                AnimateScale animateScale8 = new AnimateScale();

                CommonBehavior commonBehavior15 = new CommonBehavior();

                CommonTimeNode commonTimeNode20 = new CommonTimeNode() { Id = (UInt32Value)20U, Duration = "166", Deceleration = 50000 };

                StartConditionList startConditionList18 = new StartConditionList();
                Condition condition18 = new Condition() { Delay = "1834" };

                startConditionList18.Append(condition18);

                commonTimeNode20.Append(startConditionList18);

                TargetElement targetElement15 = new TargetElement();
                ShapeTarget shapeTarget15 = new ShapeTarget() { ShapeId = shapeId };

                targetElement15.Append(shapeTarget15);

                commonBehavior15.Append(commonTimeNode20);
                commonBehavior15.Append(targetElement15);
                ToPosition toPosition8 = new ToPosition() { X = 100000, Y = 100000 };

                animateScale8.Append(commonBehavior15);
                animateScale8.Append(toPosition8);

                childTimeNodeList5.Append(setBehavior1);
                childTimeNodeList5.Append(animateEffect1);
                childTimeNodeList5.Append(animate1);
                childTimeNodeList5.Append(animate2);
                childTimeNodeList5.Append(animate3);
                childTimeNodeList5.Append(animate4);
                childTimeNodeList5.Append(animate5);
                childTimeNodeList5.Append(animateScale1);
                childTimeNodeList5.Append(animateScale2);
                childTimeNodeList5.Append(animateScale3);
                childTimeNodeList5.Append(animateScale4);
                childTimeNodeList5.Append(animateScale5);
                childTimeNodeList5.Append(animateScale6);
                childTimeNodeList5.Append(animateScale7);
                childTimeNodeList5.Append(animateScale8);

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

                Condition condition19 = new Condition() { Event = TriggerEventValues.OnPrevious, Delay = "0" };

                TargetElement targetElement16 = new TargetElement();
                SlideTarget slideTarget1 = new SlideTarget();

                targetElement16.Append(slideTarget1);

                condition19.Append(targetElement16);

                previousConditionList1.Append(condition19);

                NextConditionList nextConditionList1 = new NextConditionList();

                Condition condition20 = new Condition() { Event = TriggerEventValues.OnNext, Delay = "0" };

                TargetElement targetElement17 = new TargetElement();
                SlideTarget slideTarget2 = new SlideTarget();

                targetElement17.Append(slideTarget2);

                condition20.Append(targetElement17);

                nextConditionList1.Append(condition20);

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
