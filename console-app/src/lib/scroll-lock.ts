export function createScrollLock(getStyle: () => {overflow:string}) {
 const owners=new Set<symbol>(); let previous='';
 return () => {
  const token=Symbol(),style=getStyle();
  if(!owners.size)previous=style.overflow;
  owners.add(token);style.overflow='hidden';
  return () => {if(!owners.delete(token))return;if(!owners.size)getStyle().overflow=previous;};
 };
}
export const lockBodyScroll=createScrollLock(()=>document.body.style);
