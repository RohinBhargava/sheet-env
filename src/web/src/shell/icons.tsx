// Small stroke glyphs used by the popup menus.
import type { CSSProperties, ReactNode } from 'react';

export interface IconProps { size?: number; style?: CSSProperties; className?: string; color?: string }

function Glyph({ size = 16, children, style, className, color }: IconProps & { children: ReactNode }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" stroke={color ?? 'currentColor'} strokeWidth={1} strokeLinecap="round" strokeLinejoin="round" style={{ display: 'block', ...style }} className={className} aria-hidden>
      {children}
    </svg>
  );
}

export const ChevronRight = (p: IconProps) => <Glyph {...p}><path d="M6 4l4 4-4 4" /></Glyph>;
export const CheckGlyph = (p: IconProps) => <Glyph {...p}><path d="M3 8.5l3.2 3L13 4.5" /></Glyph>;
export const DotGlyph = (p: IconProps) => <Glyph {...p}><circle cx="8" cy="8" r="2.2" fill="currentColor" /></Glyph>;
