export enum DayRange {
  Seven,
  Thirty,
  Sixty,
}

export interface TimeModel {
  range: DayRange;
  name: string;
}
