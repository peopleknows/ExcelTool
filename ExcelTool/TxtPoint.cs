using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExcelTool
{
    public class TxtPoint
    {
        
        [DisplayName("数据编号")]
        public int Id { get; set; } = 0;//在表中的序号
        [DisplayName("文件名称")]
        public string OverlayName { get; set; } = "UnKnown";//其实是文件名字/表的名字
        [DisplayName("车站名称")]
        public string StationName { get; set; } = "UnKnown";//应该先找到车站名称统一赋值

        [DisplayName("类型名称")]
        public string TypeName { get; set; } = "";//特殊类型的名字;保存绝缘节的名称/道岔的名称SWn/信号机的名称S-上行？X-下行
        [DisplayName("里程")]
        public double KiloPos { get; set; } = 0.0;
        [DisplayName("轨道编号")]
        public string TrackName { get; set; } = "UnKnown";// 轨道编号
        
        [DisplayName("经度(毫秒)")]
        public double Longtitude { get; set; } = 0.0;//经度
        [DisplayName("纬度(毫秒)")]
        public double Latitude { get; set; } = 0.0;//纬度
        [DisplayName("高程(厘米)")]
        public double Height { get; set; } = 0.0;
        [DisplayName("航向角(弧度)")]
        public double Bear { get; set; } = 0.0;//航向角

        [DisplayName("增量航向角(弧度)")]
        public double DeltaBear { get; set; } = 0.0;//增量航向角

        [DisplayName("备注")]
        public string Tag { get; set; } = "UnKnown";//备注：如绝缘节表有说明哪一个点是和信号机处于同一位置

        public PointType type;//点的类型，根据此编写位置
        public TxtPoint()
        {

        }
        public TxtPoint(int index, string overlayname, string stationName, string trackname, double lat, double lng, double
             height, double bear, double deltabear, double KiloPos)
        {
            this.Id = index;
            this.OverlayName = overlayname;
            this.StationName = StationName;
            this.TrackName = trackname;
            this.Latitude = lat;
            this.Longtitude = lng;
            this.Height = height;
            this.Bear = bear;
            this.DeltaBear = deltabear;
            this.KiloPos = KiloPos;
        }

    }

    public enum PointType
    {
        Unknown = 0,
        Signal = 1,
        PiecePoint = 2,
        Switch = 3,
        Joint = 4//绝缘节
    }
}
