import { DayRange, TimeModel } from "../models/chartModel";

export function getTimeRange(): TimeModel[] {
  return [
    { range: DayRange.Seven, name: "7 days" },
    { range: DayRange.Thirty, name: "30 days" },
    { range: DayRange.Sixty, name: "60 days" },
  ];
}

export function getChart1Points(range: DayRange): any[] {
  switch (range) {
    case DayRange.Seven:
      return chart1Points_7D;
    case DayRange.Thirty:
      return chart1Points_30D;
    case DayRange.Sixty:
      return chart1Points_60D;
    default:
      return [];
  }
}

export function getChart2Points(range: DayRange): any[] {
  switch (range) {
    case DayRange.Seven:
      return chart2Points_7D;
    case DayRange.Thirty:
      return chart2Points_30D;
    case DayRange.Sixty:
      return chart2Points_60D;
    default:
      return [];
  }
}

const chart1Points_7D = [
  { x: new Date("01/01"), y: 18000 },
  { x: new Date("01/06"), y: 14000 },
  { x: new Date("01/11"), y: 19000 },
  { x: new Date("01/16"), y: 13000 },
  { x: new Date("01/21"), y: 21000 },
  { x: new Date("01/26"), y: 18000 },
  { x: new Date("01/31"), y: 23000 },
];

const chart2Points_7D = [
  { x: new Date("01/01"), y: 8000 },
  { x: new Date("01/06"), y: 10000 },
  { x: new Date("01/11"), y: 100 },
  { x: new Date("01/16"), y: 9000 },
  { x: new Date("01/21"), y: 11000 },
  { x: new Date("01/26"), y: 7000 },
  { x: new Date("01/31"), y: 7200 },
];

const chart1Points_30D = [
  { x: new Date("01/01"), y: 18000 },
  { x: new Date("01/02"), y: 14000 },
  { x: new Date("01/03"), y: 19000 },
  { x: new Date("01/04"), y: 13000 },
  { x: new Date("01/05"), y: 21000 },
  { x: new Date("01/06"), y: 18000 },
  { x: new Date("01/07"), y: 18000 },
  { x: new Date("01/08"), y: 14000 },
  { x: new Date("01/09"), y: 19000 },
  { x: new Date("01/10"), y: 13000 },
  { x: new Date("01/11"), y: 21000 },
  { x: new Date("01/12"), y: 18000 },
  { x: new Date("01/13"), y: 23000 },
  { x: new Date("01/14"), y: 18000 },
  { x: new Date("01/15"), y: 14000 },
  { x: new Date("01/16"), y: 19000 },
  { x: new Date("01/17"), y: 13000 },
  { x: new Date("01/18"), y: 21000 },
  { x: new Date("01/19"), y: 18000 },
  { x: new Date("01/20"), y: 23000 },
  { x: new Date("01/21"), y: 18000 },
  { x: new Date("01/22"), y: 14000 },
  { x: new Date("01/23"), y: 19000 },
  { x: new Date("01/24"), y: 13000 },
  { x: new Date("01/25"), y: 21000 },
  { x: new Date("01/26"), y: 18000 },
  { x: new Date("01/27"), y: 23000 },
  { x: new Date("01/28"), y: 13000 },
  { x: new Date("01/29"), y: 21000 },
  { x: new Date("01/30"), y: 18000 },
  { x: new Date("01/31"), y: 23000 },
];

const chart2Points_30D = [
  { x: new Date("01/01"), y: 18000 },
  { x: new Date("01/02"), y: 18000 },
  { x: new Date("01/03"), y: 18000 },
  { x: new Date("01/04"), y: 18000 },
  { x: new Date("01/05"), y: 18000 },
  { x: new Date("01/06"), y: 14000 },
  { x: new Date("01/07"), y: 18000 },
  { x: new Date("01/08"), y: 18000 },
  { x: new Date("01/09"), y: 18000 },
  { x: new Date("01/10"), y: 18000 },
  { x: new Date("01/11"), y: 19000 },
  { x: new Date("01/12"), y: 18000 },
  { x: new Date("01/13"), y: 18000 },
  { x: new Date("01/14"), y: 18000 },
  { x: new Date("01/15"), y: 18000 },
  { x: new Date("01/16"), y: 13000 },
  { x: new Date("01/17"), y: 18000 },
  { x: new Date("01/18"), y: 18000 },
  { x: new Date("01/19"), y: 18000 },
  { x: new Date("01/20"), y: 18000 },
  { x: new Date("01/21"), y: 12000 },
  { x: new Date("01/22"), y: 18000 },
  { x: new Date("01/23"), y: 18000 },
  { x: new Date("01/24"), y: 18000 },
  { x: new Date("01/25"), y: 18000 },
  { x: new Date("01/26"), y: 14000 },
  { x: new Date("01/27"), y: 18000 },
  { x: new Date("01/28"), y: 18000 },
  { x: new Date("01/29"), y: 18000 },
  { x: new Date("01/30"), y: 18000 },
  { x: new Date("01/31"), y: 15000 },
];

const chart1Points_60D = [
  { x: new Date("01/01"), y: 18000 },
  { x: new Date("01/02"), y: 14000 },
  { x: new Date("01/03"), y: 19000 },
  { x: new Date("01/04"), y: 13000 },
  { x: new Date("01/05"), y: 21000 },
  { x: new Date("01/06"), y: 18000 },
  { x: new Date("01/07"), y: 18000 },
  { x: new Date("01/08"), y: 14000 },
  { x: new Date("01/09"), y: 19000 },
  { x: new Date("01/10"), y: 13000 },
  { x: new Date("01/11"), y: 21000 },
  { x: new Date("01/12"), y: 18000 },
  { x: new Date("01/13"), y: 23000 },
  { x: new Date("01/14"), y: 18000 },
  { x: new Date("01/15"), y: 14000 },
  { x: new Date("01/16"), y: 19000 },
  { x: new Date("01/17"), y: 13000 },
  { x: new Date("01/18"), y: 21000 },
  { x: new Date("01/19"), y: 18000 },
  { x: new Date("01/20"), y: 23000 },
  { x: new Date("01/21"), y: 18000 },
  { x: new Date("01/22"), y: 14000 },
  { x: new Date("01/23"), y: 19000 },
  { x: new Date("01/24"), y: 13000 },
  { x: new Date("01/25"), y: 21000 },
  { x: new Date("01/26"), y: 18000 },
  { x: new Date("01/27"), y: 23000 },
  { x: new Date("01/28"), y: 13000 },
  { x: new Date("01/29"), y: 21000 },
  { x: new Date("01/30"), y: 18000 },
  { x: new Date("01/31"), y: 23000 },
  { x: new Date("02/01"), y: 18000 },
  { x: new Date("02/02"), y: 14000 },
  { x: new Date("02/03"), y: 19000 },
  { x: new Date("02/04"), y: 13000 },
  { x: new Date("02/05"), y: 21000 },
  { x: new Date("02/06"), y: 18000 },
  { x: new Date("02/07"), y: 18000 },
  { x: new Date("02/08"), y: 14000 },
  { x: new Date("02/09"), y: 19000 },
  { x: new Date("02/10"), y: 13000 },
  { x: new Date("02/11"), y: 21000 },
  { x: new Date("02/12"), y: 18000 },
  { x: new Date("02/13"), y: 23000 },
  { x: new Date("02/14"), y: 18000 },
  { x: new Date("02/15"), y: 14000 },
  { x: new Date("02/16"), y: 19000 },
  { x: new Date("02/17"), y: 13000 },
  { x: new Date("02/18"), y: 21000 },
  { x: new Date("02/19"), y: 18000 },
  { x: new Date("02/20"), y: 23000 },
  { x: new Date("02/21"), y: 18000 },
  { x: new Date("02/22"), y: 14000 },
  { x: new Date("02/23"), y: 19000 },
  { x: new Date("02/24"), y: 13000 },
  { x: new Date("02/25"), y: 21000 },
  { x: new Date("02/26"), y: 18000 },
  { x: new Date("02/27"), y: 23000 },
  { x: new Date("02/28"), y: 13000 },
];

const chart2Points_60D = [
  { x: new Date("01/01"), y: 18000 },
  { x: new Date("01/02"), y: 18000 },
  { x: new Date("01/03"), y: 18000 },
  { x: new Date("01/04"), y: 18000 },
  { x: new Date("01/05"), y: 18000 },
  { x: new Date("01/06"), y: 14000 },
  { x: new Date("01/07"), y: 18000 },
  { x: new Date("01/08"), y: 18000 },
  { x: new Date("01/09"), y: 18000 },
  { x: new Date("01/10"), y: 18000 },
  { x: new Date("01/11"), y: 19000 },
  { x: new Date("01/12"), y: 18000 },
  { x: new Date("01/13"), y: 18000 },
  { x: new Date("01/14"), y: 18000 },
  { x: new Date("01/15"), y: 18000 },
  { x: new Date("01/16"), y: 13000 },
  { x: new Date("01/17"), y: 18000 },
  { x: new Date("01/18"), y: 18000 },
  { x: new Date("01/19"), y: 18000 },
  { x: new Date("01/20"), y: 18000 },
  { x: new Date("01/21"), y: 12000 },
  { x: new Date("01/22"), y: 18000 },
  { x: new Date("01/23"), y: 18000 },
  { x: new Date("01/24"), y: 18000 },
  { x: new Date("01/25"), y: 18000 },
  { x: new Date("01/26"), y: 14000 },
  { x: new Date("01/27"), y: 18000 },
  { x: new Date("01/28"), y: 18000 },
  { x: new Date("01/29"), y: 18000 },
  { x: new Date("01/30"), y: 18000 },
  { x: new Date("01/31"), y: 15000 },
  { x: new Date("02/01"), y: 15000 },
  { x: new Date("02/02"), y: 15000 },
  { x: new Date("02/03"), y: 15000 },
  { x: new Date("02/04"), y: 15000 },
  { x: new Date("02/05"), y: 15000 },
  { x: new Date("02/06"), y: 15000 },
  { x: new Date("02/07"), y: 15000 },
  { x: new Date("02/08"), y: 15000 },
  { x: new Date("02/09"), y: 15000 },
  { x: new Date("02/10"), y: 15000 },
  { x: new Date("02/11"), y: 15000 },
  { x: new Date("02/12"), y: 15000 },
  { x: new Date("02/13"), y: 15000 },
  { x: new Date("02/14"), y: 15000 },
  { x: new Date("02/15"), y: 15000 },
  { x: new Date("02/16"), y: 15000 },
  { x: new Date("02/17"), y: 15000 },
  { x: new Date("02/18"), y: 15000 },
  { x: new Date("02/19"), y: 15000 },
  { x: new Date("02/20"), y: 15000 },
  { x: new Date("02/21"), y: 15000 },
  { x: new Date("02/22"), y: 15000 },
  { x: new Date("02/23"), y: 15000 },
  { x: new Date("02/24"), y: 15000 },
  { x: new Date("02/25"), y: 15000 },
  { x: new Date("02/26"), y: 15000 },
  { x: new Date("02/27"), y: 15000 },
  { x: new Date("02/28"), y: 15000 },
];
