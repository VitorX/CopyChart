using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string fromPresentation1 = @"source.pptx";
            string toPresentation2 = @"target.pptx";
            string toPresentation2Backup = @"target - Copy.pptx";
            File.Delete(toPresentation2);
            File.Copy(toPresentation2Backup, toPresentation2);
            CopyChart(fromPresentation1, toPresentation2);

            Process.Start(toPresentation2);
        }

        private static void CopyChart(string fromPresentation1, string toPresentation2)
        {
            
            using (PresentationDocument ppt1 = PresentationDocument.Open(fromPresentation1, false))
            using (PresentationDocument ppt2 = PresentationDocument.Open(toPresentation2, true))
            {

                SlideId fromSlideId = ppt1.PresentationPart.Presentation.SlideIdList.GetFirstChild<SlideId>();
                string fromRelId = fromSlideId.RelationshipId;

                SlideId toSlideId = ppt2.PresentationPart.Presentation.SlideIdList.GetFirstChild<SlideId>();
                string toRelId = fromSlideId.RelationshipId;

                SlidePart fromSlidePart = (SlidePart)ppt1.PresentationPart.GetPartById(fromRelId);
                SlidePart toSlidePart = (SlidePart)ppt2.PresentationPart.GetPartById(fromRelId);

                var graphFrame=fromSlidePart.Slide.CommonSlideData.ShapeTree.GetFirstChild<GraphicFrame>().CloneNode(true);
                GroupShapeProperties groupShapeProperties = toSlidePart.Slide.CommonSlideData.ShapeTree.GetFirstChild<GroupShapeProperties>();
                toSlidePart.Slide.CommonSlideData.ShapeTree.InsertAfter(graphFrame, groupShapeProperties);

                ChartPart fromChartPart = fromSlidePart.ChartParts.First();
                ChartPart toChartPart = toSlidePart.AddNewPart<ChartPart>("rId2");


                using (StreamReader streamReader = new StreamReader(fromChartPart.GetStream()))
                using (StreamWriter streamWriter = new StreamWriter(toChartPart.GetStream(FileMode.Create)))
                {
                    streamWriter.Write(streamReader.ReadToEnd());
                }

                EmbeddedPackagePart fromEmbeddedPackagePart1 = fromChartPart.EmbeddedPackagePart;
                EmbeddedPackagePart toEmbeddedPackagePart1 = toChartPart.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");

                using (StreamReader streamReader = new StreamReader(fromEmbeddedPackagePart1.GetStream()))
                    toEmbeddedPackagePart1.FeedData(streamReader.BaseStream);
            
            }
        }
    }
}
