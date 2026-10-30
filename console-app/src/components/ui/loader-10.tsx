import * as React from 'react';
import { cn } from '@/lib/utils';
export interface GooeyLoaderProps extends React.HTMLAttributes<HTMLDivElement> {
 primaryColor?: string;
 secondaryColor?: string;
 borderColor?: string;
}
const GooeyLoader = React.forwardRef<HTMLDivElement, GooeyLoaderProps>(
 ({className,primaryColor,secondaryColor,borderColor,style,...props},ref) => {
  const filterId = `gooey-loader-${React.useId().replace(/[^a-zA-Z0-9_-]/g,'')}`;
  const variables = {
   '--gooey-primary-color':primaryColor || 'hsl(var(--primary))',
   '--gooey-secondary-color':secondaryColor || 'hsl(var(--secondary))',
   '--gooey-border-color':borderColor || 'hsl(var(--border))',
   '--gooey-filter':`url(#${filterId})`,...style,
  } as React.CSSProperties;
  return <div ref={ref} className={cn('relative flex items-center justify-center text-sm',className)} style={variables} role="status" aria-label="Carregando dados" {...props}>
   <svg className="absolute w-0 h-0" aria-hidden="true" focusable="false"><defs><filter id={filterId}><feGaussianBlur in="SourceGraphic" stdDeviation={12} result="blur"/><feColorMatrix in="blur" mode="matrix" values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 48 -7" result="goo"/><feComposite in="SourceGraphic" in2="goo" operator="atop"/></filter></defs></svg>
   <style>{`
    .gooey-loader{width:12em;height:3em;position:relative;overflow:hidden;border-bottom:8px solid var(--gooey-border-color);filter:var(--gooey-filter)}
    .gooey-loader::before,.gooey-loader::after{content:'';position:absolute;border-radius:50%}
    .gooey-loader::before{width:22em;height:18em;background-color:var(--gooey-primary-color);left:-2em;bottom:-18em;animation:gooey-loader-wee1 2s linear infinite}
    .gooey-loader::after{width:16em;height:12em;background-color:var(--gooey-secondary-color);left:-4em;bottom:-12em;animation:gooey-loader-wee2 2s linear infinite .75s}
    @keyframes gooey-loader-wee1{0%{transform:translateX(-10em) rotate(0deg)}100%{transform:translateX(7em) rotate(180deg)}}
    @keyframes gooey-loader-wee2{0%{transform:translateX(-8em) rotate(0deg)}100%{transform:translateX(8em) rotate(180deg)}}
    @media(prefers-reduced-motion:reduce){.gooey-loader::before,.gooey-loader::after{animation:none;width:2.8em;height:2.8em;bottom:-1.4em;left:3.5em}.gooey-loader::after{left:5.8em}}
   `}</style><div className="gooey-loader" aria-hidden="true"/>
  </div>;
 });
GooeyLoader.displayName='GooeyLoader';
export {GooeyLoader};
