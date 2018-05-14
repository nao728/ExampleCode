using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using DevExpress.XtraCharts;
using DocumentFormat.OpenXml.Packaging;
using InstanceDistributor.Util;
using PileFoundation.core.analysis;
using PileFoundation.core.pile.constructor;
using PileFoundation.core.section;
using PileFoundation.core.spring;
using PileFoundation.model.entity.analysis;
using PileFoundation.model.entity.geologic;
using PileFoundation.model.entity.member;
using PileFoundation.model.entity.member.material;
using PileFoundation.model.entity.section;
using PileFoundation.model.entity.type;
using WordEditorAPI.core.table;
using WordEditorAPI.util;
using WordPublisher.model;
using WordPublisher.util;
using PileFoundation.core.pile.design;
using PileFoundation.core.footing;
using PileFoundation.model.impl.tableRetriever;
using PileFoundation.model.entity;
using WordEditorAPI.core.table.retrieve.implementation;
using PileFoundation.model.impl.tableRetriever.rebar;
using PileFoundation.model.type;

namespace PileFoundation.model.impl.draft
{
    public class PileFoundationDrafter : Draftable
    {
        private HCPileDesigner hcpileDesigner;

        public PileFoundationDrafter(HCPileDesigner hcpileDesigner) 
        {
            this.hcpileDesigner = hcpileDesigner;
        }

        private string repositoryDir = @"D:\git\WordPublisher\WordPublisher\Test\repository\";
        private string imageDirPath = @"D:\git\WordPublisher\WordPublisher\Resources\images\";


        public bool CanWriteDraft(string chDirName)
        {
            return chDirName.Contains("16");
        }

        public void WriteDraft(WordprocessingDocument draftWordDoc)
        {
            // input : select pile construction type
            EnumPileConstructionType pileConstructionType = EnumPileConstructionType.HCP;

            //ExampleForSteelPileOnly(draftWordDoc);

            ExampleForHCPile(draftWordDoc);

        }

        private void ExampleForHCPile(WordprocessingDocument draftWordDoc)
        {
            
            ////////////////////////////////////////////////////////////////////////////////
            // local variables
            SteelPileConstructor steelPileConstructor = hcpileDesigner.steelPileConstructor;
            PHCPileConstructor phcPileConstructor = hcpileDesigner.phcPileConstructor;

            SteelPileMember steelPileMember = hcpileDesigner.steelPileMember;
            PileSection phcPileSection = hcpileDesigner.phcPileSection;

            PileModelConfiguration pileModelConfiguration = hcpileDesigner.pileModelConfiguration;

            PileSection steelPileSection = hcpileDesigner.steelPileSection;
            SpringCoefficientEstimater coeffEstimater = hcpileDesigner.coeffEstimater;

            SteelPileFEA pileFea_fixVer = hcpileDesigner.simPileFea_fixVer;
            SteelPileFEA pileFea_fixHor = hcpileDesigner.simPileFea_fixHor;
            SteelPileFEA pileFea_fixBend = hcpileDesigner.simPileFea_fixBend;

            SteelPileFEA pileFea_fixVerEQ = hcpileDesigner.simPileFea_fixVerEQ;
            SteelPileFEA pileFea_fixHorEQ = hcpileDesigner.simPileFea_fixHorEQ;
            SteelPileFEA pileFea_fixBendEQ = hcpileDesigner.simPileFea_fixBendEQ;

            SteelPileFEA pileFea_hingeVer = hcpileDesigner.simPileFea_hingeVer;
            SteelPileFEA pileFea_hingeHor = hcpileDesigner.simPileFea_hingeHor;

            SteelPileFEA pileFea_hingeVerEQ = hcpileDesigner.simPileFea_hingeVerEQ;
            SteelPileFEA pileFea_hingeHorEQ = hcpileDesigner.simPileFea_hingeHorEQ;


            PileStabilityDesigner pileStabilityDesigner = hcpileDesigner.pileStabilityDesigner;
            PileStabilityDesigner pileStabilityDesignerEQ = hcpileDesigner.pileStabilityDesignerEQ;


            SteelPileDesigner steelPileDesigner_fixVer  =hcpileDesigner.steelPileDesigner_fixVer;
            SteelPileDesigner steelPileDesigner_fixHor  =hcpileDesigner.steelPileDesigner_fixHor;
            SteelPileDesigner steelPileDesigner_fixBend  =hcpileDesigner.steelPileDesigner_fixBend;
            SteelPileDesigner steelPileDesigner_hingeHor =hcpileDesigner.steelPileDesigner_hingeHor; 
            SteelPileDesigner steelPileDesigner_hingeVer =hcpileDesigner.steelPileDesigner_hingeVer ;
            
                                                            
            PHCPileDesigner phcPileDesigner_fixVer =hcpileDesigner.phcPileDesigner_fixVer;
            PHCPileDesigner phcPileDesigner_fixHor =hcpileDesigner.phcPileDesigner_fixHor;
            PHCPileDesigner phcPileDesigner_fixBend =hcpileDesigner.phcPileDesigner_fixBend;
            PHCPileDesigner phcPileDesigner_hingeHor =hcpileDesigner.phcPileDesigner_hingeHor;
            PHCPileDesigner phcPileDesigner_hingeVer =hcpileDesigner.phcPileDesigner_hingeVer;


            FootingBodyDesigner footingBodyDesigner = hcpileDesigner.footingBodyDesigner;
            FootingDesigner footingDesigner = hcpileDesigner.footingDesigner;
            RebarEmbedmentDesigner rebarEmbedmentDesigner = hcpileDesigner.rebarEmbedmentDesigner;



            GeologicalColumn phcPile_gc = hcpileDesigner.optPHCPile_gc;
            GeologicalColumn steelPile_gc = hcpileDesigner.optSteelPile_gc;
            
            // 
            

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_두부보강철근의정착길이_이형철근의허용응력", 0, new RebarMaterialTableRetriever(hcpileDesigner.footingDesigner.rebarMember));

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_두부보강철근의정착길이_이형철근의표준치수", 0, new RebarDimensionTableRetriever(hcpileDesigner.footingDesigner.rebarMember));

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_두부보강철근의정착길이_콘크리트의허용부착응력", 0, new FootingMaterialTableRetriever(hcpileDesigner.footingDesigner.footingMaterial));


            TableStamper.StampContentsOnly(draftWordDoc, "테이블_휨압축응력계수휨인장응력계수", 1, new PhiCalcTableRetriever(hcpileDesigner));

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝머리의작용력", 2, new PilePeakLoadTableRetriever(hcpileDesigner));


            TableStamper.StampContentsOnly(draftWordDoc, "테이블_본체해석결과_상시", 2, 
                new PileLoadTableRetriever(hcpileDesigner.load_fixVer, hcpileDesigner.load_fixHor, hcpileDesigner.load_fixBend, hcpileDesigner.load_hingeVer, hcpileDesigner.load_hingeHor)
                );
            TableStamper.StampContentsOnly(draftWordDoc, "테이블_본체해석결과_지진시", 2,
                new PileLoadTableRetriever(hcpileDesigner.load_fixVerEQ, hcpileDesigner.load_fixHorEQ, hcpileDesigner.load_fixBendEQ, hcpileDesigner.load_hingeVerEQ, hcpileDesigner.load_hingeHorEQ)
                );
            
            
            
            TableEraser.EraseTable(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터이하");

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터초과", 3,
                new SpringConstantTableRetriever(steelPileMember, steelPileSection, coeffEstimater, pileFea_fixVer));

            FillResultTable(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리고정_연직력최대", pileFea_fixVer.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리고정_수평력최대", pileFea_fixHor.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리고정_모멘트최대", pileFea_fixBend.ResultTable);

            FillResultTable(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리힌지_연직력최대", pileFea_hingeHor.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리힌지_수평력최대", pileFea_hingeVer.ResultTable);

            FillResultTable(draftWordDoc, "테이블_말뚝의해석_지진시_말뚝머리고정_연직력최대", pileFea_fixVerEQ.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_지진시_말뚝머리고정_수평력최대", pileFea_fixHorEQ.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_지진시_말뚝머리고정_모멘트최대", pileFea_fixBendEQ.ResultTable);

            FillResultTable(draftWordDoc, "테이블_말뚝의해석_지진시_말뚝머리힌지_연직력최대", pileFea_hingeHorEQ.ResultTable);
            FillResultTable(draftWordDoc, "테이블_말뚝의해석_지진시_말뚝머리힌지_수평력최대", pileFea_hingeVerEQ.ResultTable);



            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_1_PileCapServiceabilityfixed_Vmax.jpg",
                pileFea_fixVer.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );
            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_2_PileCapServiceabilityfixed_Hmax.jpg",
                pileFea_fixHor.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_3_PileCapServiceabilityfixed_Mmax.jpg",
                pileFea_fixBend.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_4_PileCapServiceabilityhinge_Vmax.jpg",
                pileFea_hingeVer.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_5_PileCapServiceabilityhinge_Hmax.jpg",
                pileFea_hingeHor.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_6_PileCapSeismicfixed_Vmax.jpg",
                pileFea_fixVerEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );
            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_7_PileCapSeismicfixed_Hmax.jpg",
                pileFea_fixHorEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_8_PileCapSeismicfixed_Mmax.jpg",
                pileFea_fixBendEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "PileCapSeismichinge_Vmax.jpg",
                pileFea_hingeVerEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_8_PileCapSeismicfixed_Mmax.jpg",
                pileFea_fixBendEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_9_PileCapSeismichinge_Vmax.jpg",
                pileFea_hingeVerEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(966 / 2, 519 / 2, imageDirPath + "16_10_PileCapSeismichinge_Hmax.jpg",
                pileFea_hingeHorEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                })
            );

            PileChartPlotter.PlotSingleChart(1798/2, 1420/2, imageDirPath + "16_11_HCP_connecting.jpg",
                new List<SeriesPoint>[]{

                pileFea_fixVer.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_fixHor.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_fixBend.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_hingeVer.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_hingeHor.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),

                pileFea_fixVerEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_fixHorEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_fixBendEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_hingeVerEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),
                pileFea_hingeHorEQ.ResultTable.RecordList.ConvertAll<SeriesPoint>((rec) =>
                {
                    return new SeriesPoint(rec.Depth, rec.Moment);
                }),

                }
            , -400, 400);



            FillBendingStressTable(draftWordDoc, "테이블_말뚝본체의응력검토_상시_휨응력", steelPileDesigner_fixBend.StressTable, phcPileDesigner_fixBend.StressTable);
            FillShearStressTable(draftWordDoc, "테이블_말뚝본체의응력검토_상시_전단응력", steelPileDesigner_fixBend.StressTable, phcPileDesigner_fixBend.StressTable);

            FillBendingStressTable(draftWordDoc, "테이블_말뚝본체의응력검토_지진시_휨응력", steelPileDesigner_fixBend.StressTableEQ, phcPileDesigner_fixBend.StressTableEQ);
            FillShearStressTable(draftWordDoc, "테이블_말뚝본체의응력검토_지진시_전단응력", steelPileDesigner_fixBend.StressTableEQ, phcPileDesigner_fixBend.StressTableEQ);


            List<SeriesPoint> seriesPointList_fsa = new List<SeriesPoint>();
            steelPileDesigner_fixBend.StressTable.RecordList.ForEach(rec => {
                seriesPointList_fsa.Add(new SeriesPoint(rec.Depth, rec.AllowableBendingStress));
            });
            List<SeriesPoint> seriesPointList_fpa = new List<SeriesPoint>();
            phcPileDesigner_fixBend.StressTable.RecordList.ForEach(rec => {
                seriesPointList_fpa.Add(new SeriesPoint(rec.Depth, rec.AllowableBendingStress));
            });

            PileChartPlotter.PlotSingleChart(811 / 2, 621 / 2, imageDirPath + "16_12_Serviceability_f.jpg",
                new List<SeriesPoint>[]{
                steelPileDesigner_fixBend.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixBend.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_fixHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_fixVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_hingeVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_hingeVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                
                steelPileDesigner_hingeHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_hingeHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                seriesPointList_fsa,
                seriesPointList_fpa

                }
            , 0, 150, 35);




            List<SeriesPoint> seriesPointList_tausa = new List<SeriesPoint>();
            steelPileDesigner_fixBend.StressTable.RecordList.ForEach(rec =>
            {
                seriesPointList_tausa.Add(new SeriesPoint(rec.Depth, rec.AllowableShearStress));
            });
            List<SeriesPoint> seriesPointList_taupa = new List<SeriesPoint>();
            phcPileDesigner_fixBend.StressTable.RecordList.ForEach(rec =>
            {
                seriesPointList_taupa.Add(new SeriesPoint(rec.Depth, rec.AllowableShearStress));
            });

            PileChartPlotter.PlotSingleChart(811 / 2, 621 / 2, imageDirPath + "16_13_Serviceability_tau.jpg",
                new List<SeriesPoint>[]{
                steelPileDesigner_fixBend.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixBend.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_fixHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_fixVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_hingeVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_hingeVer.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                
                steelPileDesigner_hingeHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_hingeHor.StressTable.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                seriesPointList_tausa,
                seriesPointList_taupa
                }
            , 0, 90, 20);



            List<SeriesPoint> seriesPointList_fsaEQ = new List<SeriesPoint>();
            steelPileDesigner_fixBend.StressTableEQ.RecordList.ForEach(rec =>
            {
                seriesPointList_fsaEQ.Add(new SeriesPoint(rec.Depth, rec.AllowableBendingStress));
            });
            List<SeriesPoint> seriesPointList_fpaEQ = new List<SeriesPoint>();
            phcPileDesigner_fixBend.StressTableEQ.RecordList.ForEach(rec =>
            {
                seriesPointList_fpaEQ.Add(new SeriesPoint(rec.Depth, rec.AllowableBendingStress));
            });

            PileChartPlotter.PlotSingleChart(811 / 2, 621 / 2, imageDirPath + "16_14_Seismic_f.jpg",
                new List<SeriesPoint>[]{
                steelPileDesigner_fixBend.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixBend.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_fixHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_fixVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_fixVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                steelPileDesigner_hingeVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_hingeVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                
                steelPileDesigner_hingeHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),
                phcPileDesigner_hingeHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.BendingStress);
                }),

                seriesPointList_fsaEQ,
                seriesPointList_fpaEQ,
 
                }
            , 0, 220, 70);





            List<SeriesPoint> seriesPointList_tausaEQ = new List<SeriesPoint>();
            steelPileDesigner_fixBend.StressTableEQ.RecordList.ForEach(rec =>
            {
                seriesPointList_tausaEQ.Add(new SeriesPoint(rec.Depth, rec.AllowableShearStress));
            });
            List<SeriesPoint> seriesPointList_taupaEQ = new List<SeriesPoint>();
            phcPileDesigner_fixBend.StressTableEQ.RecordList.ForEach(rec =>
            {
                seriesPointList_taupaEQ.Add(new SeriesPoint(rec.Depth, rec.AllowableShearStress));
            });


            PileChartPlotter.PlotSingleChart(811 / 2, 621 / 2, imageDirPath + "16_15_Seismic_tau.jpg",
                new List<SeriesPoint>[]{
                steelPileDesigner_fixBend.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixBend.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_fixHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_fixVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_fixVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                steelPileDesigner_hingeVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_hingeVer.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                
                steelPileDesigner_hingeHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),
                phcPileDesigner_hingeHor.StressTableEQ.RecordList.ConvertAll<SeriesPoint>((rec)=>{
                    return new SeriesPoint(rec.Depth, rec.ShearStress);
                }),

                seriesPointList_tausaEQ,
                seriesPointList_taupaEQ
                }
            , 0, 130, 40);


            List<SeriesPoint> dataset = new List<SeriesPoint>();
            dataset.Add(new SeriesPoint("연직방향", hcpileDesigner.redundantRatioAxial));
            dataset.Add(new SeriesPoint("말뚝재료 축하중", hcpileDesigner.redundantRatioVertical));
            dataset.Add(new SeriesPoint("수평변위", hcpileDesigner.redundantRatioDeflection));
            dataset.Add(new SeriesPoint("휨응력", hcpileDesigner.redundantRatioBending));
            dataset.Add(new SeriesPoint("전단응력", hcpileDesigner.redundantRatioShear));

            PileChartPlotter.PlotSingleBarChart(1451 / 2, 654 / 2, imageDirPath + "16_16_result.jpg", dataset, 100, 20);



            TableStamper.StampContentsOnly(draftWordDoc, "테이블_결과정리및여용력검토_결과정리", 1,
                new PileResultTableRetriever(hcpileDesigner));



            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////
            // Serialize Json
            JsonDeliverer.Deliver<SteelPileConstructor>(repositoryDir, () => steelPileConstructor);
            JsonDeliverer.Deliver<PileSection>(repositoryDir, () => steelPileSection);
            JsonDeliverer.Deliver<PHCPileConstructor>(repositoryDir, () => phcPileConstructor);
            JsonDeliverer.Deliver<PileSection>(repositoryDir, () => phcPileSection);

            JsonDeliverer.Deliver<PileModelConfiguration>(repositoryDir, () => pileModelConfiguration);

            JsonDeliverer.Deliver<GeologicalColumn>(repositoryDir, () => steelPile_gc);
            JsonDeliverer.Deliver<GeologicalColumn>(repositoryDir, () => phcPile_gc);

            JsonDeliverer.Deliver<SpringCoefficientEstimater>(repositoryDir, () => coeffEstimater);

            JsonDeliverer.Deliver<SteelPileDesigner>(repositoryDir, () => steelPileDesigner_fixVer);
            JsonDeliverer.Deliver<SteelPileDesigner>(repositoryDir, () => steelPileDesigner_fixHor);
            JsonDeliverer.Deliver<SteelPileDesigner>(repositoryDir, () => steelPileDesigner_fixBend);

            JsonDeliverer.Deliver<PHCPileDesigner>(repositoryDir, () => phcPileDesigner_fixVer);
            JsonDeliverer.Deliver<PHCPileDesigner>(repositoryDir, () => phcPileDesigner_fixHor);
            JsonDeliverer.Deliver<PHCPileDesigner>(repositoryDir, () => phcPileDesigner_fixBend);


            JsonDeliverer.Deliver<FootingBodyDesigner>(repositoryDir, () => footingBodyDesigner);
            JsonDeliverer.Deliver<FootingDesigner>(repositoryDir, () => footingDesigner);
            JsonDeliverer.Deliver<RebarEmbedmentDesigner>(repositoryDir, () => rebarEmbedmentDesigner);



            JsonDeliverer.Deliver<PileStabilityDesigner>(repositoryDir, () => pileStabilityDesigner);
            JsonDeliverer.Deliver<PileStabilityDesigner>(repositoryDir, () => pileStabilityDesignerEQ);
            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////

        }

        /**
        private void ExampleForSteelPileOnly(WordprocessingDocument draftWordDoc)
        {
            EnumConstructionMethodType constructionMethodType = EnumConstructionMethodType.PRD;

            ///////////////////////////////////////////////////////////////////////////
            // case : Steel Pile Only
            ///////////////////////////////////////////////////////////////////////////

            SteelPileConstructor steelPileConstructor = new SteelPileConstructor();

            // input : textbox diameter
            double input_textDiameter = 0.5;

            // query : diamter
            double[] steelDiameterSets = SteelPileConstructor.QuerySteelPileDiameterSets(input_textDiameter);

            // input : select diameter
            int input_selectDiameter = 0;
            double diameter = steelDiameterSets[input_selectDiameter];

            // query : thickness
            double[] steelThicknessSets = SteelPileConstructor.QueryThicknessSets(diameter);

            // input : select thickness
            int input_selectThickness = 1;

            double thickness = steelThicknessSets[input_selectThickness];

            // input : textbox thickessCorrosion
            double thickessCorrosion = 0.002;

            // input : textbox Es
            double E_s = 200000;

            // input : select steel type
            EnumSteelType steelType = EnumSteelType.SKK400;

            // input : textbox length
            double length = 8.42;

            SteelPileMember member = steelPileConstructor.CreateSteelPileMember(diameter, thickness, thickessCorrosion, length, E_s, steelType);

            PileSection steelPileSection = SectionalPropertyCalculatorAdapter.Calculate(member.Dimension);

            JsonDeliverer.Deliver<SteelPileConstructor>(repositoryDir, () => steelPileConstructor);
            JsonDeliverer.Deliver<PileSection>(repositoryDir, () => steelPileSection);

            // input : table
            double[] zCoords = PileMeshGenerator.MeshNodeZ(0, member.Dimension.Length, 0.5);

            GeologicalColumn gc = new GeologicalColumn();
            for (int i = 0; i < zCoords.Length; i++)
            {
                gc.AddStratum(zCoords[i], 10, EnumGroundType.SANDY_SOIL);
            }

            double[,] forces = new double[,]
            {
                {173.703, 0, -557.776, 0, -154.187, 0},
                {173.710, 0, -549.670, 0, -153.272, 0},
                {173.703, 0, -557.776, 0, -154.187, 0},

                {88.275, 0, -352.856, 0, -94.091, 0},
                {88.275, 0, -352.856, 0, -94.091, 0},
                {10.560, 0, -352.082, 0, -178.251, 0},
            };
            SpringCoefficientEstimater coeffEstimater = new SpringCoefficientEstimater(constructionMethodType);
            coeffEstimater.EstimateCoefficient();


            PileAnalysisModeler modeler = new PileAnalysisModeler(coeffEstimater, member, steelPileSection, gc, forces);

            modeler.CreateSteelPileAnalysisModel();
            modeler.Solve();

            // post-process
            string forcePath = Environment.CurrentDirectory + @"\force_1_LoadToPile.bin";
            BinaryReader binReader = new BinaryReader(new FileStream(forcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));

            int nElements = modeler.getNElements();
            for (int i = 0; i < nElements; i++)
            {
                double iDx = binReader.ReadDouble();
                double iDy = binReader.ReadDouble();
                double iDz = binReader.ReadDouble();
                double iMx = binReader.ReadDouble();
                double iMy = binReader.ReadDouble();
                double iMz = binReader.ReadDouble();

                double jDx = binReader.ReadDouble();
                double jDy = binReader.ReadDouble();
                double jDz = binReader.ReadDouble();
                double jMx = binReader.ReadDouble();
                double jMy = binReader.ReadDouble();
                double jMz = binReader.ReadDouble();

                Debug.Print("{0}, {1}, {2}, {3}, {4}, {5}", iDx, iDy, iDz, iMx, iMy, iMz);
                Debug.Print("{0}, {1}, {2}, {3}, {4}, {5}\n", jDx, jDy, jDz, jMx, jMy, jMz);
            }

            modeler.ReadOutput();

            

            /////////////////////////// table
            if (member.Dimension.Length >= 10)
            {
                #region 테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터초과

                TableEraser.EraseTable(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터이하");

                TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터초과", 3, (iRow, iCol) =>
                {
                    HorizontalSpringConstantCalculator horSpringCalculator = new HorizontalSpringConstantCalculator(member, steelPileSection, 1.0, 0.5, 3.0, coeffEstimater);
                    HorizontalSpringConstantCalculator horSpringCalculator_eq = new HorizontalSpringConstantCalculator(member, steelPileSection, 2.0, 0.5, 2.0, coeffEstimater);

                    string val = null;

                    if (iRow == 0)
                    {
                        switch (iCol)
                        {
                            case 6:
                                val = gc.GetStratum(0).N.ToString();
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.GetStratum(0).GroundType);
                                break;
                        }
                    }
                    else if (iRow >= 2 && iRow <= 18)
                    {
                        int index = iRow - 1;

                        horSpringCalculator.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);
                        horSpringCalculator_eq.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);

                        switch (iCol)
                        {
                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;

                            case 7:
                                val = NumericFormatter.Format(horSpringCalculator.E_0, 0, true);
                                break;

                            case 8:
                                val = NumericFormatter.Format(horSpringCalculator.K_h, 0, true);
                                break;

                            case 9:
                                val = NumericFormatter.Format(horSpringCalculator_eq.K_h, 0, true);
                                break;

                            case 10:
                                val = NumericFormatter.Format(horSpringCalculator.f_h, 3, true);
                                break;

                            case 11:
                                val = NumericFormatter.Format(horSpringCalculator_eq.f_h, 3, true);
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.StratumList[index].GroundType);
                                break;
                        }
                    }
                    else if (iRow == 23)
                    {
                        int index = gc.StratumList.Count - 1;
                        horSpringCalculator.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);
                        horSpringCalculator_eq.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);

                        switch (iCol)
                        {
                            case 4:
                                val = (index + 1).ToString();
                                break;

                            case 5:
                                val = NumericFormatter.Format(Math.Abs(gc.StratumList[index].Depth), 2, true);
                                break;

                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;
                        }
                    }
                    else if (iRow == 22)
                    {
                        int index = gc.StratumList.Count - 2;
                        horSpringCalculator.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);
                        horSpringCalculator_eq.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);

                        switch (iCol)
                        {
                            case 4:
                                val = (index + 1).ToString();
                                break;

                            case 5:
                                val = NumericFormatter.Format(Math.Abs(gc.StratumList[index].Depth), 2, true);
                                break;

                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;

                            case 7:
                                val = NumericFormatter.Format(horSpringCalculator.E_0, 0, true);
                                break;

                            case 8:
                                val = NumericFormatter.Format(horSpringCalculator.K_h, 0, true);
                                break;

                            case 9:
                                val = NumericFormatter.Format(horSpringCalculator_eq.K_h, 0, true);
                                break;

                            case 10:
                                val = NumericFormatter.Format(horSpringCalculator.f_h, 3, true);
                                break;

                            case 11:
                                val = NumericFormatter.Format(horSpringCalculator_eq.f_h, 3, true);
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.StratumList[index].GroundType);
                                break;
                        }
                    }
                    else if (iRow == 21)
                    {
                        int index = gc.StratumList.Count - 3;
                        horSpringCalculator.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);
                        horSpringCalculator_eq.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);

                        switch (iCol)
                        {
                            case 4:
                                val = (index + 1).ToString();
                                break;

                            case 5:
                                val = NumericFormatter.Format(Math.Abs(gc.StratumList[index].Depth), 2, true);
                                break;

                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;

                            case 7:
                                val = NumericFormatter.Format(horSpringCalculator.E_0, 0, true);
                                break;

                            case 8:
                                val = NumericFormatter.Format(horSpringCalculator.K_h, 0, true);
                                break;

                            case 9:
                                val = NumericFormatter.Format(horSpringCalculator_eq.K_h, 0, true);
                                break;

                            case 10:
                                val = NumericFormatter.Format(horSpringCalculator.f_h, 3, true);
                                break;

                            case 11:
                                val = NumericFormatter.Format(horSpringCalculator_eq.f_h, 3, true);
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.StratumList[index].GroundType);
                                break;
                        }
                    }

                    
                    return val;
                });

                #endregion 테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터초과
            }
            else
            {
                #region 테이블_말뚝모델링_수평방향스프링정수및주변마찰력_10미터미만

                TableEraser.EraseTable(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터초과");

                int[] delRow = new int[21 - gc.StratumList.Count + 1];
                for (int i = 0; i < delRow.Length; i++)
                {
                    delRow[i] = gc.StratumList.Count + i;
                }
                TableEraser.EraseTableRow(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터이하", delRow);

                TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝모델링_수평방향스프링정수및주면마찰력_10미터이하", 3, (iRow, iCol) =>
                {
                    HorizontalSpringConstantCalculator horSpringCalculator = new HorizontalSpringConstantCalculator(member, steelPileSection, 1.0, 0.5, 3.0, coeffEstimater);
                    HorizontalSpringConstantCalculator horSpringCalculator_eq = new HorizontalSpringConstantCalculator(member, steelPileSection, 2.0, 0.5, 2.0, coeffEstimater);

                    string val = null;

                    if (iRow == 0)
                    {
                        switch (iCol)
                        {
                            case 6:
                                val = gc.GetStratum(0).N.ToString();
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.GetStratum(0).GroundType);
                                break;
                        }
                    }
                    else if (iRow - 1 < gc.StratumList.Count - 1)
                    {
                        int index = iRow - 1;

                        horSpringCalculator.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);
                        horSpringCalculator_eq.Calculate(gc.StratumList[index].GroundType, gc.StratumList[index].N);

                        switch (iCol)
                        {
                            case 4:
                                val = (index + 1).ToString();
                                break;

                            case 5:
                                val = NumericFormatter.Format(Math.Abs(gc.StratumList[index].Depth), 2, true);
                                break;

                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;

                            case 7:
                                val = NumericFormatter.Format(horSpringCalculator.E_0, 0, true);
                                break;

                            case 8:
                                val = NumericFormatter.Format(horSpringCalculator.K_h, 0, true);
                                break;

                            case 9:
                                val = NumericFormatter.Format(horSpringCalculator_eq.K_h, 0, true);
                                break;

                            case 10:
                                val = NumericFormatter.Format(horSpringCalculator.f_h, 3, true);
                                break;

                            case 11:
                                val = NumericFormatter.Format(horSpringCalculator_eq.f_h, 3, true);
                                break;

                            case 12:
                                val = EnumUtil.GetEnum<EnumGroundType>(gc.StratumList[index].GroundType);
                                break;
                        }
                    }
                    else if (iRow - 1 == gc.StratumList.Count - 1)
                    {
                        int index = iRow - 1;

                        switch (iCol)
                        {
                            case 4:
                                val = (index + 1).ToString();
                                break;

                            case 5:
                                val = NumericFormatter.Format(Math.Abs(gc.StratumList[index].Depth), 2, true);
                                break;

                            case 6:
                                val = gc.StratumList[index].N.ToString();
                                break;
                        }
                    }

                    return val;
                });

                #endregion 테이블_말뚝모델링_수평방향스프링정수및주변마찰력_10미터미만
            }

            #region 테이블_말뚝의해석_상시_말뚝머리고정_연직력최대

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리고정_연직력최대", 1, (iRow, iCol) =>
                {
                    string val = null;

                    AnalysisResultTable table = modeler.GetAnaylsisResultTables()[0];

                    if (iRow <= 15)
                    {
                        switch (iCol)
                        {
                            case 0:
                                val = NumericFormatter.Format(table.RecordList[iRow].Depth, 2, true);
                                break;

                            case 1:
                                val = NumericFormatter.Format(table.RecordList[iRow].Displacement, 6, true);
                                break;

                            case 2:
                                val = NumericFormatter.Format(table.RecordList[iRow].Axial, 3, true);
                                break;

                            case 3:
                                val = NumericFormatter.Format(table.RecordList[iRow].Shear, 3, true);
                                break;

                            case 4:
                                val = NumericFormatter.Format(table.RecordList[iRow].Moment, 3, true);
                                break;
                        }
                    }

                    return val;
                });

            #endregion 테이블_말뚝의해석_상시_말뚝머리고정_연직력최대

            #region 테이블_말뚝의해석_상시_말뚝머리고정_수평력최대

            TableStamper.StampContentsOnly(draftWordDoc, "테이블_말뚝의해석_상시_말뚝머리고정_수평력최대", 1, (iRow, iCol) =>
            {
                string val = null;

                AnalysisResultTable table = modeler.GetAnaylsisResultTables()[1];

                if (iRow <= 15)
                {
                    switch (iCol)
                    {
                        case 0:
                            val = NumericFormatter.Format(table.RecordList[iRow].Depth, 2, true);
                            break;

                        case 1:
                            val = NumericFormatter.Format(table.RecordList[iRow].Displacement, 6, true);
                            break;

                        case 2:
                            val = NumericFormatter.Format(table.RecordList[iRow].Axial, 3, true);
                            break;

                        case 3:
                            val = NumericFormatter.Format(table.RecordList[iRow].Shear, 3, true);
                            break;

                        case 4:
                            val = NumericFormatter.Format(table.RecordList[iRow].Moment, 3, true);
                            break;
                    }
                }

                return val;
            });

            #endregion 테이블_말뚝의해석_상시_말뚝머리고정_수평력최대
        }
        **/

        private void FillBendingStressTable(WordprocessingDocument draftWordDoc, string tableBookmarkName, PileStressTable steelStressTable, PileStressTable phcStressTable)
        {
            TableStamper.StampContentsOnly(draftWordDoc, tableBookmarkName, 2, (iRow, iCol) => {
                string val = null;

                if (iRow < 16) {
                    if (iRow < steelStressTable.RecordList.Count - 1)
                    {
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].Depth, 2, false);
                                break;
                            case 2:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].BendingStress, 2, false);
                                break;
                            case 3:
                                val = "-";
                                break;
                            case 4:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].AllowableBendingStress, 2, false);
                                break;
                            case 5:
                                val = "-";
                                break;
                            case 6:
                                val = steelStressTable.RecordList[iRow].IsOKBending ? "O.K" : "N.G";
                                break;
                        }
                    }
                    else if (iRow == steelStressTable.RecordList.Count - 1)
                    {
                        int index = iRow - (steelStressTable.RecordList.Count - 1);
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].Depth, 2, false);
                                break;
                            case 2:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].BendingStress, 2, false);
                                break;
                            case 3:
                                val = NumericFormatter.Format(phcStressTable.RecordList[0].BendingStress, 2, false);
                                break;
                            case 4:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].AllowableBendingStress, 2, false);
                                break;
                            case 5:
                                val = NumericFormatter.Format(phcStressTable.RecordList[0].AllowableBendingStress, 2, false);
                                break;
                            case 6:
                                val = steelStressTable.RecordList[iRow].IsOKBending && phcStressTable.RecordList[0].IsOKBending ? "O.K" : "N.G";
                                break;
                        }
                    } 
                    else if (iRow > steelStressTable.RecordList.Count - 1)
                    {
                        int index = iRow - (steelStressTable.RecordList.Count - 1);
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].Depth, 2, false);
                                break;
                            case 2:
                                val = "-";
                                break;
                            case 3:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].BendingStress, 2, false);
                                break;
                            case 4:
                                val = "-";
                                break;
                            case 5:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].AllowableBendingStress, 2, false);
                                break;
                            case 6:
                                val = phcStressTable.RecordList[index].IsOKBending ? "O.K" : "N.G";
                                break;
                        }
                    }
                }
                else if (iRow == 17)
                {
                    int index = phcStressTable.RecordList.Count - 1;
                    switch (iCol)
                    {
                        case 1:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].Depth, 2, false);
                            break;
                        case 2:
                            val = "-";
                            break;
                        case 3:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].BendingStress, 2, false);
                            break;
                        case 4:
                            val = "-";
                            break;
                        case 5:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].AllowableBendingStress, 2, false);
                            break;
                        case 6:
                            val = phcStressTable.RecordList[index].IsOKBending ? "O.K" : "N.G";
                            break;
                    }
                }

                //if (val == null) 
                //{
                //    val = iRow + "/" + iCol;
                //}

                return val;
            });
        }

        

        private void FillShearStressTable(WordprocessingDocument draftWordDoc, string tableBookmarkName, PileStressTable steelStressTable, PileStressTable phcStressTable)
        {
            TableStamper.StampContentsOnly(draftWordDoc, tableBookmarkName, 2, (iRow, iCol) =>
            {
                string val = null;

                if (iRow < 16)
                {
                    if (iRow < steelStressTable.RecordList.Count - 1)
                    {
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].Depth, 2, false);
                                break;
                            case 2:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].ShearStress, 2, false);
                                break;
                            case 3:
                                val = "-";
                                break;
                            case 4:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].AllowableShearStress, 2, false);
                                break;
                            case 5:
                                val = "-";
                                break;
                            case 6:
                                val = steelStressTable.RecordList[iRow].IsOKShear ? "O.K" : "N.G";
                                break;
                        }
                    }
                    else if (iRow == steelStressTable.RecordList.Count - 1)
                    {
                        int index = iRow - (steelStressTable.RecordList.Count - 1);
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].Depth, 2, false);
                                break;
                            case 2:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].ShearStress, 2, false);
                                break;
                            case 3:
                                val = NumericFormatter.Format(phcStressTable.RecordList[0].ShearStress, 2, false);
                                break;
                            case 4:
                                val = NumericFormatter.Format(steelStressTable.RecordList[iRow].AllowableShearStress, 2, false);
                                break;
                            case 5:
                                val = NumericFormatter.Format(phcStressTable.RecordList[0].AllowableShearStress, 2, false);
                                break;
                            case 6:
                                val = steelStressTable.RecordList[iRow].IsOKBending && phcStressTable.RecordList[0].IsOKShear ? "O.K" : "N.G";
                                break;
                        }
                    }
                    else if (iRow > steelStressTable.RecordList.Count - 1)
                    {
                        int index = iRow - (steelStressTable.RecordList.Count - 1);
                        switch (iCol)
                        {
                            case 1:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].Depth, 2, false);
                                break;
                            case 2:
                                val = "-";
                                break;
                            case 3:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].ShearStress, 2, false);
                                break;
                            case 4:
                                val = "-";
                                break;
                            case 5:
                                val = NumericFormatter.Format(phcStressTable.RecordList[index].AllowableShearStress, 2, false);
                                break;
                            case 6:
                                val = phcStressTable.RecordList[index].IsOKShear ? "O.K" : "N.G";
                                break;
                        }
                    }
                }
                else if (iRow == 17)
                {
                    int index = phcStressTable.RecordList.Count - 1;
                    switch (iCol)
                    {
                        case 1:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].Depth, 2, false);
                            break;
                        case 2:
                            val = "-";
                            break;
                        case 3:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].ShearStress, 2, false);
                            break;
                        case 4:
                            val = "-";
                            break;
                        case 5:
                            val = NumericFormatter.Format(phcStressTable.RecordList[index].AllowableShearStress, 2, false);
                            break;
                        case 6:
                            val = phcStressTable.RecordList[index].IsOKShear ? "O.K" : "N.G";
                            break;
                    }
                }

                //if (val == null) 
                //{
                //    val = iRow + "/" + iCol;
                //}

                return val;
            });
        }

        private void FillResultTable(WordprocessingDocument draftWordDoc, string tableBookmarkName, AnalysisResultTable table)
        {
            TableStamper.StampContentsOnly(draftWordDoc, tableBookmarkName, 1, (iRow, iCol) =>
            {
                string val = null;

                //AnalysisResultTable table = modeler.GetAnaylsisResultTables()[1];

                if (iRow <= 15)
                {
                    switch (iCol)
                    {
                        case 0:
                            val = NumericFormatter.Format(table.RecordList[iRow].Depth, 2, true);
                            break;

                        case 1:
                            val = NumericFormatter.Format(table.RecordList[iRow].Displacement, 6, true);
                            break;

                        case 2:
                            val = NumericFormatter.Format(table.RecordList[iRow].Axial, 3, true);
                            break;

                        case 3:
                            val = NumericFormatter.Format(table.RecordList[iRow].Shear, 3, true);
                            break;

                        case 4:
                            val = NumericFormatter.Format(table.RecordList[iRow].Moment, 3, true);
                            break;
                    }
                }
                else if (iRow == 17)
                {
                    int lastIndex = table.RecordList.Count - 1;
                    switch (iCol)
                    {
                        case 0:
                            val = NumericFormatter.Format(table.RecordList[lastIndex].Depth, 2, true);
                            break;

                        case 1:
                            val = NumericFormatter.Format(table.RecordList[lastIndex].Displacement, 6, true);
                            break;

                        case 2:
                            val = NumericFormatter.Format(table.RecordList[lastIndex].Axial, 3, true);
                            break;

                        case 3:
                            val = NumericFormatter.Format(table.RecordList[lastIndex].Shear, 3, true);
                            break;

                        case 4:
                            val = NumericFormatter.Format(table.RecordList[lastIndex].Moment, 3, true);
                            break;
                    }
                }


                return val;
            });
        }
    }
}