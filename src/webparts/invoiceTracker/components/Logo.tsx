import * as React from "react";

const Logo: React.FC<{ height: number; width: number; isNavOpen: boolean }> = ({
  height,
  width,
  isNavOpen,
}) => {
  return (
    <svg
      width={width}
      height={height}
      viewBox="0 0 192 192"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      {/* Your SVG paths here */}
      <path d="M100.235 12.0821C61.0686 ..." fill="#FFC000" />
      {/* rest of paths */}
    </svg>
  );
};

export default Logo;
