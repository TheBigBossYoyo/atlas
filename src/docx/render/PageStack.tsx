import React from 'react';
import type { Document } from '../model/document';
import type { Page } from '../layout/pageTypes';
import type { Theme } from '../parser/theme';
import { PageView } from './PageView';
import './__styles__/page-view.css';

export type PageStackProps = {
  pages: ReadonlyArray<Page>;
  zoom: number;
  document: Document;
  theme?: Theme;
};

export const PageStack: React.FC<PageStackProps> = ({ pages, zoom, document, theme }) => {
  return (
    <div className="docx-page-stack">
      {pages.map((page, idx) => (
        <PageView
          key={`page-${page.sectionIndex}-${page.pageIndex}-${idx}`}
          page={page}
          zoom={zoom}
          document={document}
          theme={theme}
        />
      ))}
    </div>
  );
};
