/** Values produced by the frozen Overview engine, without a second calculation. */
export interface OverviewHomeSnapshot {
 real:boolean;
 totals:{rev:number;netPftOpt:number;conv:number;cac:number;cost:number};
 targets:{revenue:number;cpa:number;budget:number};
 days:{date:string;revenue:number}[];
}
export function homeOverviewValues(snapshot:OverviewHomeSnapshot|null){
 if(!snapshot?.real)return null;
 return {sales:snapshot.totals.rev,profit:snapshot.totals.netPftOpt,orders:snapshot.totals.conv,media:snapshot.totals.cost,cac:snapshot.totals.cac};
}
