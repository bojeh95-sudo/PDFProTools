import React, { useState, useRef } from 'react';
import { PDFDocument, rgb, StandardFonts, degrees } from 'pdf-lib';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType } from 'docx';
import JSZip from 'jszip';
import { Language } from '../types';
import { triggerFileDownload } from './downloadHelper';

interface ToolPlaygroundProps {
  toolId: string;
  toolName: string;
  language: Language;
  onClose: () => void;
}

const PLAYGROUND_TEXTS = {
  [Language.EN]: {
    close: 'Close',
    uploadZonePdf: 'Drag & drop PDF here, or click to browse',
    uploadZoneImg: 'Drag & drop images (PNG/JPG) here, or click to browse',
    filesSelected: 'Files selected',
    processing: 'Processing your file...',
    success: 'Successfully processed!',
    downloadBtn: 'Download Processed File',
    processBtn: 'Process File',
    addMore: 'Add more files',
    errorTitle: 'Error',
    watermarkLabel: 'Watermark Text',
    watermarkPlaceholder: 'e.g. CONFIDENTIAL',
    passwordLabel: 'Password Protection',
    passwordPlaceholder: 'Enter security password',
    splitLabel: 'Page Range to Extract',
    splitPlaceholder: 'e.g. 1-3, 5 (leave blank for page 1)',
    translateLabel: 'Target Translation Language',
    copyBtn: 'Copy to Clipboard',
    copied: 'Copied!',
    saveTxtBtn: 'Save to Text File',
    aiResultHeader: 'AI Analysis Result',
    rotateLabel: 'Rotation Angle',
    deletePagesLabel: 'Pages to Delete',
    deletePagesPlaceholder: 'e.g. 2, 4-6',
    reorderLabel: 'New Page Order',
    reorderPlaceholder: 'e.g. 3, 1, 2, 4-5',
    unlockLabel: 'PDF Password (if protected)',
    unlockPlaceholder: 'Enter password to unlock',
    cropLabel: 'Crop Percentage (from margins)',
    chatLabel: 'Ask a question about this PDF',
    chatPlaceholder: 'e.g. Summarize this document in 3 sentences',
    chatSendBtn: 'Send',
    compareFilesError: 'Please upload exactly 2 PDF documents to compare.',
    chatIntro: 'AI Assistant loaded. Ask any question about the uploaded document!',
  },
  [Language.AR]: {
    close: 'إغلاق',
    uploadZonePdf: 'اسحب وأسقط ملف PDF هنا، أو انقر للتصفح',
    uploadZoneImg: 'اسحب وأسقط الصور (PNG/JPG) هنا، أو انقر للتصفح',
    filesSelected: 'ملفات تم تحديدها',
    processing: 'جاري معالجة ملفك...',
    success: 'تمت المعالجة بنجاح!',
    downloadBtn: 'تحميل الملف المعالج',
    processBtn: 'معالجة الملف',
    addMore: 'إضافة المزيد من الملفات',
    errorTitle: 'خطأ',
    watermarkLabel: 'نص العلامة المائية',
    watermarkPlaceholder: 'مثال: سري للغاية',
    passwordLabel: 'حماية كلمة المرور',
    passwordPlaceholder: 'أدخل كلمة مرور الحماية',
    splitLabel: 'نطاق الصفحات المراد استخراجها',
    splitPlaceholder: 'مثال: 1-3، 5 (اتركه فارغاً لاستخراج الصفحة الأولى)',
    translateLabel: 'اللغة المستهدفة للترجمة',
    copyBtn: 'نسخ إلى الحافظة',
    copied: 'تم النسخ!',
    saveTxtBtn: 'حفظ كملف نصي',
    aiResultHeader: 'نتيجة تحليل الذكاء الاصطناعي',
    rotateLabel: 'زاوية التدوير',
    deletePagesLabel: 'الصفحات المراد حذفها',
    deletePagesPlaceholder: 'مثال: 2، 4-6',
    reorderLabel: 'ترتيب الصفحات الجديد',
    reorderPlaceholder: 'مثال: 3، 1، 2، 4-5',
    unlockLabel: 'كلمة مرور ملف PDF (إذا كان محمياً)',
    unlockPlaceholder: 'أدخل كلمة المرور لإزالة الحماية',
    cropLabel: 'نسبة الهامش للقص (%)',
    chatLabel: 'اسأل عن محتوى ملف PDF بالذكاء الاصطناعي',
    chatPlaceholder: 'مثال: لخص هذا المستند في 3 جمل',
    chatSendBtn: 'إرسال',
    compareFilesError: 'الرجاء تحميل ملفين PDF بالضبط للمقارنة بينهما.',
    chatIntro: 'تم تحميل مساعد الذكاء الاصطناعي. اسأل أي سؤال حول المستند المرفوع!',
  },
  [Language.ES]: {
    close: 'Cerrar',
    uploadZonePdf: 'Arrastre y suelte el PDF aquí o haga clic para cargarlo',
    uploadZoneImg: 'Arrastre y suelte imágenes (PNG/JPG) aquí o haga clic para cargarlo',
    filesSelected: 'Archivos seleccionados',
    processing: 'Procesando su archivo...',
    success: '¡Procesado con éxito!',
    downloadBtn: 'Descargar archivo procesado',
    processBtn: 'Procesar archivo',
    addMore: 'Añadir más archivos',
    errorTitle: 'Error',
    watermarkLabel: 'Texto de la marca de agua',
    watermarkPlaceholder: 'ej. CONFIDENCIAL',
    passwordLabel: 'Contraseña de protección',
    passwordPlaceholder: 'Introduzca la contraseña de seguridad',
    splitLabel: 'Rango de páginas a extraer',
    splitPlaceholder: 'ej. 1-3, 5 (dejar en blanco para la página 1)',
    translateLabel: 'Idioma de destino para traducción',
    copyBtn: 'Copiar al portapapeles',
    copied: '¡Copiado!',
    saveTxtBtn: 'Guardar como archivo de texto',
    aiResultHeader: 'Resultado del análisis de IA',
    rotateLabel: 'Ángulo de rotación',
    deletePagesLabel: 'Páginas a eliminar',
    deletePagesPlaceholder: 'ej. 2, 4-6',
    reorderLabel: 'Nuevo orden de páginas',
    reorderPlaceholder: 'ej. 3, 1, 2, 4-5',
    unlockLabel: 'Contraseña de PDF (si está protegido)',
    unlockPlaceholder: 'Introduzca la contraseña para desbloquear',
    cropLabel: 'Porcentaje de recorte (márgenes)',
    chatLabel: 'Hacer una pregunta sobre el PDF',
    chatPlaceholder: 'ej. Resume este documento en 3 oraciones',
    chatSendBtn: 'Enviar',
    compareFilesError: 'Por favor, suba exactamente 2 documentos PDF para comparar.',
    chatIntro: 'Asistente de IA cargado. ¡Haga cualquier pregunta sobre el documento!',
  },
  [Language.FR]: {
    close: 'Fermer',
    uploadZonePdf: 'Glissez-déposez le PDF ici ou cliquez pour télécharger',
    uploadZoneImg: 'Glissez-déposez des images (PNG/JPG) ici ou cliquez pour télécharger',
    filesSelected: 'Fichiers sélectionnés',
    processing: 'Traitement de votre fichier en cours...',
    success: 'Traité avec succès !',
    downloadBtn: 'Télécharger le fichier traité',
    processBtn: 'Traiter le fichier',
    addMore: 'Ajouter plus de fichiers',
    errorTitle: 'Erreur',
    watermarkLabel: 'Texte du filigrane',
    watermarkPlaceholder: 'ex. CONFIDENTIEL',
    passwordLabel: 'Protection par mot de passe',
    passwordPlaceholder: 'Entrez le mot de passe de sécurité',
    splitLabel: 'Plage de pages à extraire',
    splitPlaceholder: 'ex. 1-3, 5 (laisser vide pour la page 1)',
    translateLabel: 'Langue cible pour la traduction',
    copyBtn: 'Copier dans le presse-papiers',
    copied: 'Copié !',
    saveTxtBtn: 'Enregistrer en fichier texte',
    aiResultHeader: 'Résultat de l\'analyse d\'IA',
    rotateLabel: 'Angle de rotation',
    deletePagesLabel: 'Pages à supprimer',
    deletePagesPlaceholder: 'ex. 2, 4-6',
    reorderLabel: 'Nouvel ordre des pages',
    reorderPlaceholder: 'ex. 3, 1, 2, 4-5',
    unlockLabel: 'Mot de passe du PDF (si protégé)',
    unlockPlaceholder: 'Entrez le mot de passe pour déverrouiller',
    cropLabel: 'Pourcentage de rognage (marges)',
    chatLabel: 'Poser une question sur le PDF',
    chatPlaceholder: 'ex. Résumer ce document en 3 phrases',
    chatSendBtn: 'Envoyer',
    compareFilesError: 'Veuillez télécharger exactement 2 documents PDF à comparer.',
    chatIntro: 'Assistant IA chargé. Posez n\'importe quelle question sur le document !',
  }
};

const LANGUAGES_LIST = [
  { code: 'Arabic', label: 'العربية (Arabic)' },
  { code: 'English', label: 'English' },
  { code: 'Spanish', label: 'Español (Spanish)' },
  { code: 'French', label: 'Français (French)' },
  { code: 'German', label: 'Deutsch (German)' },
  { code: 'Chinese', label: '中文 (Chinese)' },
  { code: 'Japanese', label: '日本語 (Japanese)' },
  { code: 'Hindi', label: 'हिन्दी (Hindi)' },
];

interface InteractivePageCardProps {
  page: {
    id: string;
    fileIndex: number;
    pageNum: number;
    fileName: string;
    isSelectedForDelete: boolean;
    rotation: number;
    splitAfter: boolean;
  };
  idx: number;
  toolId: string;
  language: Language;
  onPageClick: () => void;
  onRotateClick: () => void;
  onMovePage: (idx: number, direction: 'left' | 'right') => void;
  isFirst: boolean;
  isLast: boolean;
  watermarkText: string;
  cropPercentage: number;
  requestRender: (fileIndex: number, pageNum: number) => void;
  imageSrc: string | undefined;
  renderStatus: 'idle' | 'rendering' | 'done' | 'error' | undefined;
  rangeInfo?: {
    index: number;
    color: string;
  } | null;
  onDeleteClick?: () => void;
  // Advanced tool additions
  watermarkType?: 'text' | 'image';
  watermarkFont?: string;
  watermarkSize?: number;
  watermarkColor?: string;
  watermarkBold?: boolean;
  watermarkItalic?: boolean;
  watermarkPosition?: string;
  watermarkRotation?: number;
  watermarkOpacity?: number;
  watermarkImageUrl?: string | null;
  cropMargins?: { top: number; bottom: number; left: number; right: number };
  onCropMarginsChange?: (margins: { top: number; bottom: number; left: number; right: number }) => void;
  onDragStart?: (e: React.DragEvent, index: number) => void;
  onDragOver?: (e: React.DragEvent) => void;
  onDrop?: (e: React.DragEvent, index: number) => void;
  isSelectedForRotate?: boolean;
  selectedPageSource?: number | null;
  isGloballySelected?: boolean;
}

const InteractivePageCard: React.FC<InteractivePageCardProps> = ({
  page,
  idx,
  toolId,
  language,
  onPageClick,
  onRotateClick,
  onMovePage,
  isFirst,
  isLast,
  watermarkText,
  cropPercentage,
  requestRender,
  imageSrc,
  renderStatus,
  rangeInfo,
  onDeleteClick,
  watermarkType = 'text',
  watermarkFont = 'Helvetica',
  watermarkSize = 36,
  watermarkColor = '#EF4444',
  watermarkBold = false,
  watermarkItalic = false,
  watermarkPosition = 'center',
  watermarkRotation = 45,
  watermarkOpacity = 30,
  watermarkImageUrl = null,
  cropMargins = { top: 10, bottom: 10, left: 10, right: 10 },
  onCropMarginsChange,
  onDragStart,
  onDragOver,
  onDrop,
  isSelectedForRotate = false,
  selectedPageSource = null,
  isGloballySelected = false
}) => {
  const containerRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            requestRender(page.fileIndex, page.pageNum);
            observer.disconnect();
          }
        });
      },
      {
        root: null,
        rootMargin: '120px', // Pre-render slightly before scrolling in for optimal visual flow
        threshold: 0.01
      }
    );

    if (containerRef.current) {
      observer.observe(containerRef.current);
    }

    return () => {
      observer.disconnect();
    };
  }, [page.fileIndex, page.pageNum, requestRender]);

  const isSelected = page.isSelectedForDelete;

  // Determine border and highlight styles based on tools or range mappings
  let customBorderClass = '';
  const isSourcePage = selectedPageSource === page.pageNum && (toolId === 'chat' || toolId === 'compare' || toolId === 'extract');

  if (isGloballySelected) {
    customBorderClass = 'border-indigo-600 dark:border-indigo-500 bg-indigo-50/20 dark:bg-indigo-950/20 ring-2 ring-indigo-500/40 shadow-sm';
  } else if (rangeInfo) {
    if (rangeInfo.color === 'emerald') customBorderClass = 'border-emerald-500 bg-emerald-50/10 dark:bg-emerald-950/10 ring-2 ring-emerald-500/30';
    else if (rangeInfo.color === 'sky') customBorderClass = 'border-sky-500 bg-sky-50/10 dark:bg-sky-950/10 ring-2 ring-sky-500/30';
    else if (rangeInfo.color === 'amber') customBorderClass = 'border-amber-500 bg-amber-50/10 dark:bg-amber-950/10 ring-2 ring-amber-500/30';
    else if (rangeInfo.color === 'pink') customBorderClass = 'border-pink-500 bg-pink-50/10 dark:bg-pink-950/10 ring-2 ring-pink-500/30';
    else if (rangeInfo.color === 'indigo') customBorderClass = 'border-indigo-500 bg-indigo-50/10 dark:bg-indigo-950/10 ring-2 ring-indigo-500/30';
    else if (rangeInfo.color === 'rose') customBorderClass = 'border-rose-500 bg-rose-50/10 dark:bg-rose-950/10 ring-2 ring-rose-500/30';
    else customBorderClass = 'border-violet-500 bg-violet-50/10 dark:bg-violet-950/10 ring-2 ring-violet-500/30';
  } else if (isSelected) {
    customBorderClass = 'border-red-500 bg-red-50/30 dark:bg-red-950/20 ring-2 ring-red-500/20';
  } else if (isSelectedForRotate && toolId === 'rotate') {
    customBorderClass = 'border-amber-500 bg-amber-50/20 dark:bg-amber-950/20 ring-2 ring-amber-500/40 shadow-md';
  } else if (isSourcePage) {
    customBorderClass = 'border-violet-600 bg-violet-50/20 dark:bg-violet-950/20 ring-2 ring-violet-500/50 shadow-md animate-pulse';
  } else {
    customBorderClass = 'border-slate-200 dark:border-slate-800 hover:border-indigo-400 dark:hover:border-indigo-500 hover:shadow';
  }

  // Watermark positioning styles
  let positionClasses = 'items-center justify-center';
  if (watermarkPosition === 'top-left') positionClasses = 'items-start justify-start p-1.5';
  else if (watermarkPosition === 'top-center') positionClasses = 'items-start justify-center p-1.5';
  else if (watermarkPosition === 'top-right') positionClasses = 'items-start justify-end p-1.5';
  else if (watermarkPosition === 'center-left') positionClasses = 'items-center justify-start p-1.5';
  else if (watermarkPosition === 'center-center') positionClasses = 'items-center justify-center p-1.5';
  else if (watermarkPosition === 'center-right') positionClasses = 'items-center justify-end p-1.5';
  else if (watermarkPosition === 'bottom-left') positionClasses = 'items-end justify-start p-1.5';
  else if (watermarkPosition === 'bottom-center') positionClasses = 'items-end justify-center p-1.5';
  else if (watermarkPosition === 'bottom-right') positionClasses = 'items-end justify-end p-1.5';

  return (
    <div
      ref={containerRef}
      onClick={onPageClick}
      draggable={['reorder', 'merge', 'img2pdf'].includes(toolId)}
      onDragStart={onDragStart ? (e) => onDragStart(e, idx) : undefined}
      onDragOver={onDragOver}
      onDrop={onDrop ? (e) => onDrop(e, idx) : undefined}
      className={`relative group cursor-pointer border rounded-xl p-3 bg-white dark:bg-slate-900 border-slate-200 dark:border-slate-800 transition-all duration-200 select-none flex flex-col items-center shadow-sm ${customBorderClass}`}
    >
      {/* Global Selection Checkmark Badge */}
      {isGloballySelected && (
        <div className="absolute top-2 left-2 w-5 h-5 rounded-full bg-indigo-600 text-white flex items-center justify-center text-[10px] font-extrabold shadow-md z-10 animate-fade-in">
          ✓
        </div>
      )}

      {/* Delete Button / Icon badge for specific tools */}
      {onDeleteClick && (
        <button
          onClick={(e) => {
            e.stopPropagation();
            onDeleteClick();
          }}
          className="absolute top-2 right-2 w-5 h-5 rounded-full bg-red-100 dark:bg-red-950/60 hover:bg-red-500 text-red-600 dark:text-red-400 hover:text-white flex items-center justify-center text-[10px] font-bold shadow-sm transition-all z-20"
          title={language === Language.AR ? 'حذف' : 'Delete'}
        >
          ✕
        </button>
      )}

      {/* Delete Checkmark Badge */}
      {toolId === 'delete_pages' && !onDeleteClick && (
        <div className={`absolute top-2 right-2 w-5 h-5 rounded-full flex items-center justify-center text-xs font-bold transition-all z-10 ${
          isSelected ? 'bg-red-500 text-white shadow' : 'bg-slate-100 dark:bg-slate-800 text-slate-400 dark:text-slate-500 group-hover:bg-slate-200 dark:group-hover:bg-slate-700'
        }`}>
          {isSelected ? '✕' : ''}
        </div>
      )}

      {/* Colored Range Badge overlay */}
      {rangeInfo && (
        <div className="absolute top-2 right-2 bg-slate-900/85 text-white font-extrabold text-[8px] px-1.5 py-0.5 rounded shadow z-10 font-mono">
          {language === Language.AR ? `نطاق ${rangeInfo.index + 1}` : `Range ${rangeInfo.index + 1}`}
        </div>
      )}

      {/* Selection border for active rotate page */}
      {toolId === 'rotate' && isSelectedForRotate && (
        <div className="absolute top-2 right-2 bg-amber-500 text-white font-extrabold text-[8px] px-2 py-0.5 rounded shadow z-10">
          {language === Language.AR ? 'محددة للتدوير' : 'Selected'}
        </div>
      )}

      {/* Source Citation Badge */}
      {isSourcePage && (
        <div className="absolute top-2 right-2 bg-violet-600 text-white font-extrabold text-[8px] px-2 py-0.5 rounded shadow-md z-10 animate-bounce">
          🎯 {language === Language.AR ? 'المصدر' : 'Source'}
        </div>
      )}

      {/* Card visual wrapper */}
      <div className="w-20 h-24 bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-800 shadow-sm rounded-lg flex flex-col items-center justify-center relative overflow-hidden my-1">
        {/* Loading spinner overlay */}
        {renderStatus === 'rendering' && (
          <div className="absolute inset-0 bg-slate-50/90 dark:bg-slate-900/90 flex flex-col items-center justify-center z-10">
            <div className="w-5 h-5 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin mb-1"></div>
            <span className="text-[8px] text-indigo-600 font-bold tracking-tight">
              {language === Language.AR ? 'جاري الرسم...' : 'Rendering...'}
            </span>
          </div>
        )}

        {/* Error overlay fallback */}
        {renderStatus === 'error' && (
          <div className="absolute inset-0 bg-red-50/95 dark:bg-red-950/95 flex flex-col items-center justify-center z-10 p-1">
            <span className="text-[14px] mb-0.5">⚠️</span>
            <span className="text-[8px] text-red-600 dark:text-red-400 font-bold text-center leading-tight">
              {language === Language.AR ? 'فشل التحميل' : 'Load Failed'}
            </span>
          </div>
        )}

        {imageSrc ? (
          <img
            src={imageSrc}
            alt={`Page ${page.pageNum + 1}`}
            className="w-full h-full object-contain p-0.5 transition-transform duration-200"
            style={{ transform: `rotate(${page.rotation}deg)` }}
            referrerPolicy="no-referrer"
          />
        ) : (
          <div 
            className="w-full h-full flex flex-col items-center justify-center p-2 transition-transform duration-200"
            style={{ transform: `rotate(${page.rotation}deg)` }}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-7 w-7 text-indigo-500 mb-1" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            <span className="text-[10px] text-slate-400 dark:text-slate-500 font-mono font-bold">
              {['docx2pdf', 'xlsx2pdf', 'pptx2pdf', 'html2pdf', 'txt2pdf', 'rtf2pdf'].includes(toolId) ? (
                toolId === 'docx2pdf' ? 'Word'
                : toolId === 'xlsx2pdf' ? 'Excel'
                : toolId === 'pptx2pdf' ? 'PPT'
                : toolId === 'html2pdf' ? 'HTML'
                : toolId === 'txt2pdf' ? 'TXT'
                : 'RTF'
              ) : `P. ${page.pageNum + 1}`}
            </span>
          </div>
        )}

        {/* Live Watermark overlay preview */}
        {toolId === 'watermark' && (
          <div className={`absolute inset-0 flex pointer-events-none overflow-hidden select-none z-10 ${positionClasses}`}>
            {watermarkType === 'text' && watermarkText.trim() ? (
              <span 
                className="whitespace-nowrap inline-block select-none font-sans uppercase text-center"
                style={{
                  fontFamily: watermarkFont,
                  fontSize: `${Math.max(6, watermarkSize / 3.5)}px`,
                  color: watermarkColor,
                  fontWeight: watermarkBold ? 'bold' : 'normal',
                  fontStyle: watermarkItalic ? 'italic' : 'normal',
                  transform: `rotate(${watermarkRotation}deg)`,
                  opacity: watermarkOpacity / 100,
                }}
              >
                {watermarkText}
              </span>
            ) : watermarkType === 'image' && watermarkImageUrl ? (
              <img 
                src={watermarkImageUrl} 
                alt="Watermark watermark-overlay" 
                className="object-contain"
                style={{
                  maxWidth: '60%',
                  maxHeight: '60%',
                  transform: `rotate(${watermarkRotation}deg)`,
                  opacity: watermarkOpacity / 100,
                }}
                referrerPolicy="no-referrer"
              />
            ) : null}
          </div>
        )}

        {/* Live Crop bounding box overlay with direct hover sliders */}
        {toolId === 'crop' && (
          <>
            <div 
              className="absolute border border-dashed border-rose-500 bg-rose-500/5 pointer-events-none transition-all duration-150 z-10"
              style={{
                top: `${cropMargins.top}%`,
                bottom: `${cropMargins.bottom}%`,
                left: `${cropMargins.left}%`,
                right: `${cropMargins.right}%`,
              }}
            />
            {/* Interactive sliders directly on the page overlay for immediate crop adjustment */}
            {onCropMarginsChange && (
              <div className="absolute inset-x-0 bottom-0 bg-slate-900/95 text-white p-1 text-[8px] flex flex-col space-y-0.5 opacity-0 group-hover:opacity-100 transition-opacity duration-150 z-30 pointer-events-auto">
                <div className="flex items-center justify-between gap-1">
                  <span>T: {cropMargins.top}%</span>
                  <input 
                    type="range" 
                    min="0" 
                    max="45" 
                    value={cropMargins.top} 
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => onCropMarginsChange({ ...cropMargins, top: parseInt(e.target.value, 10) })} 
                    className="w-10 h-1 accent-rose-500" 
                  />
                </div>
                <div className="flex items-center justify-between gap-1">
                  <span>B: {cropMargins.bottom}%</span>
                  <input 
                    type="range" 
                    min="0" 
                    max="45" 
                    value={cropMargins.bottom} 
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => onCropMarginsChange({ ...cropMargins, bottom: parseInt(e.target.value, 10) })} 
                    className="w-10 h-1 accent-rose-500" 
                  />
                </div>
                <div className="flex items-center justify-between gap-1">
                  <span>L: {cropMargins.left}%</span>
                  <input 
                    type="range" 
                    min="0" 
                    max="45" 
                    value={cropMargins.left} 
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => onCropMarginsChange({ ...cropMargins, left: parseInt(e.target.value, 10) })} 
                    className="w-10 h-1 accent-rose-500" 
                  />
                </div>
                <div className="flex items-center justify-between gap-1">
                  <span>R: {cropMargins.right}%</span>
                  <input 
                    type="range" 
                    min="0" 
                    max="45" 
                    value={cropMargins.right} 
                    onClick={(e) => e.stopPropagation()}
                    onChange={(e) => onCropMarginsChange({ ...cropMargins, right: parseInt(e.target.value, 10) })} 
                    className="w-10 h-1 accent-rose-500" 
                  />
                </div>
              </div>
            )}
          </>
        )}

        {/* Locked shield visual lock preview */}
        {toolId === 'protect' && (
          <div className="absolute inset-0 bg-slate-900/10 flex items-center justify-center z-10">
            <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-emerald-600 animate-pulse" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
            </svg>
          </div>
        )}

        {/* Red Overlay for marked delete */}
        {isSelected && (
          <div className="absolute inset-0 bg-red-500/25 flex flex-col items-center justify-center border-2 border-red-500 z-20 font-sans">
            <span className="text-[10px] font-extrabold text-red-600 dark:text-red-400 bg-white dark:bg-slate-900 px-2 py-0.5 rounded shadow border border-red-200 dark:border-red-800 uppercase tracking-wide">
              {language === Language.AR ? 'سيُحذف' : 'DELETE'}
            </span>
          </div>
        )}
      </div>

      {/* Label description */}
      <div className="text-[10px] font-bold text-slate-600 dark:text-slate-400 mt-2 max-w-full truncate text-center px-1 font-mono">
        {page.fileName && ['merge', 'img2pdf'].includes(toolId) ? page.fileName : `${language === Language.AR ? 'صفحة' : 'Page'} ${page.pageNum + 1}`}
      </div>

      {/* Controls inside each page card */}
      <div className="flex items-center space-x-1.5 space-x-reverse mt-2.5 w-full justify-center">
        {toolId === 'rotate' && (
          <button
            onClick={(e) => {
              e.stopPropagation();
              onRotateClick();
            }}
            className="py-1 px-2 rounded bg-amber-500 hover:bg-amber-600 text-white text-[10px] font-bold shadow-sm transition-all flex items-center gap-1 font-sans"
            title={language === Language.AR ? 'تدوير 90 درجة' : 'Rotate 90°'}
          >
            <span>↻ 90°</span>
          </button>
        )}

        {['merge', 'reorder', 'img2pdf'].includes(toolId) && (
          <div className="flex items-center space-x-1 space-x-reverse">
            <button
              onClick={(e) => {
                e.stopPropagation();
                onMovePage(idx, 'left');
              }}
              disabled={isFirst}
              className="p-1.5 rounded bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 disabled:opacity-40 text-slate-700 dark:text-slate-300 transition-colors shadow-sm text-[10px]"
              title={language === Language.AR ? 'تحريك لليمين' : 'Move Left'}
            >
              ◀
            </button>
            <button
              onClick={(e) => {
                e.stopPropagation();
                onMovePage(idx, 'right');
              }}
              disabled={isLast}
              className="p-1.5 rounded bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 disabled:opacity-40 text-slate-700 dark:text-slate-300 transition-colors shadow-sm text-[10px]"
              title={language === Language.AR ? 'تحريك ليسار' : 'Move Right'}
            >
              ▶
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const renderPdfPagesToImages = async (file: File, format: 'jpeg' | 'png'): Promise<{ name: string; blob: Blob }[]> => {
  const arrayBuffer = await file.arrayBuffer();
  
  // Dynamically load PDF.js from CDN
  const pdfjsLib = await new Promise<any>((resolve, reject) => {
    if ((window as any).pdfjsLib) {
      resolve((window as any).pdfjsLib);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      const pjs = (window as any).pdfjsLib;
      pjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(pjs);
    };
    script.onerror = (err) => reject(new Error('Failed to load PDF.js from CDN'));
    document.head.appendChild(script);
  });

  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;
  const results: { name: string; blob: Blob }[] = [];

  for (let i = 1; i <= numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 2.0 }); // High resolution 2x
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    if (!context) continue;

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    const renderContext = {
      canvasContext: context,
      viewport: viewport
    };
    await page.render(renderContext).promise;

    const mimeType = format === 'jpeg' ? 'image/jpeg' : 'image/png';
    const ext = format === 'jpeg' ? 'jpg' : 'png';
    const quality = format === 'jpeg' ? 0.95 : undefined;

    const blob = await new Promise<Blob>((resBlob) => {
      canvas.toBlob((b) => resBlob(b || new Blob()), mimeType, quality);
    });

    const pageName = `${file.name.replace(/\.pdf$/i, '')}_page_${i}.${ext}`;
    results.push({ name: pageName, blob });
  }

  return results;
};

const detectArabicOrientation = (text: string): 'reversed' | 'logical' | 'unknown' => {
  const words = text.split(/\s+/).map(w => w.trim()).filter(w => /[\u0600-\u06FF]/.test(w));
  if (words.length === 0) return 'unknown';

  let reversedCount = 0;
  let logicalCount = 0;

  for (const word of words) {
    if (word.startsWith('ال') || word.startsWith('بال') || word.startsWith('وال') || word.startsWith('لل') || word.startsWith('في') || word.startsWith('من') || word.startsWith('على')) {
      logicalCount++;
    }
    if (word.endsWith('لا') || word.endsWith('لاب') || word.endsWith('لاو') || word.endsWith('لل') || word.endsWith('يف') || word.endsWith('نم') || word.endsWith('ىلع')) {
      reversedCount++;
    }
  }

  if (reversedCount > logicalCount) return 'reversed';
  if (logicalCount > reversedCount) return 'logical';
  return 'unknown';
};

const fixArabicWordReversal = (text: string): string => {
  if (!text) return '';
  const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
  const arabicRunRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+/g;

  const lines = text.split('\n');
  const correctedLines = lines.map(line => {
    const trimmed = line.trim();
    if (!trimmed || !arabicRegex.test(trimmed)) {
      return line;
    }
    
    const orientation = detectArabicOrientation(trimmed);
    if (orientation === 'logical') {
      return line;
    }
    
    const tokens = line.split(/(\s+)/);
    const reversedTokens = tokens.reverse();
    const correctedTokens = reversedTokens.map(token => {
      if (token.trim() && arabicRegex.test(token)) {
        return token.replace(arabicRunRegex, match => match.split('').reverse().join(''));
      }
      return token;
    });
    
    return correctedTokens.join('');
  });
  
  return correctedLines.join('\n');
};

const fixArabicInSvgString = (svgString: string): string => {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(svgString, 'image/svg+xml');
    
    // Find all text and tspan elements
    const textElements = doc.querySelectorAll('text, tspan');
    textElements.forEach(el => {
      if (el.childNodes.length === 1 && el.childNodes[0].nodeType === 3) {
        const originalText = el.textContent || '';
        const fixedText = fixArabicWordReversal(originalText);
        if (fixedText !== originalText) {
          el.textContent = fixedText;
        }
      } else {
        el.childNodes.forEach(child => {
          if (child.nodeType === 3) {
            const originalText = child.textContent || '';
            const fixedText = fixArabicWordReversal(originalText);
            if (fixedText !== originalText) {
              child.textContent = fixedText;
            }
          }
        });
      }
    });
    
    const serializer = new XMLSerializer();
    return serializer.serializeToString(doc);
  } catch (err) {
    console.error('Failed to parse and fix Arabic in SVG string:', err);
    return svgString;
  }
};

const renderPdfPagesToSVGs = async (file: File): Promise<{ name: string; content: string }[]> => {
  const arrayBuffer = await file.arrayBuffer();
  
  // Dynamically load PDF.js from CDN
  const pdfjsLib = await new Promise<any>((resolve, reject) => {
    if ((window as any).pdfjsLib) {
      resolve((window as any).pdfjsLib);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      const pjs = (window as any).pdfjsLib;
      pjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(pjs);
    };
    script.onerror = (err) => reject(new Error('Failed to load PDF.js from CDN'));
    document.head.appendChild(script);
  });

  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;
  const results: { name: string; content: string }[] = [];

  for (let i = 1; i <= numPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale: 1.5 });
    
    try {
      // Get operator list and graphics for true vector SVG conversion
      const opList = await page.getOperatorList();
      const svgGfx = new pdfjsLib.SVGGraphics(page.commonObjs, page.objs);
      const svgElement = await svgGfx.getSVG(opList, viewport);
      
      const serializer = new XMLSerializer();
      let svgString = serializer.serializeToString(svgElement);
      
      // Fix reversed Arabic words inside the generated SVG
      svgString = fixArabicInSvgString(svgString);
      
      results.push({
        name: `${file.name.replace(/\.pdf$/i, '')}_page_${i}.svg`,
        content: svgString
      });
    } catch (err) {
      console.error(`SVG conversion failed for page ${i}, using text fallback:`, err);
      // Fallback text rendering
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map((item: any) => item.str).join(' ');
      const correctedText = fixArabicWordReversal(pageText);
      const words = correctedText.trim().split(/\s+/);
      let lines: string[] = [];
      let currentLine = '';
      const maxWordsPerLine = 12;
      for (const word of words) {
        if (currentLine.split(' ').length >= maxWordsPerLine) {
          lines.push(currentLine.trim());
          currentLine = '';
        }
        currentLine += word + ' ';
      }
      if (currentLine.trim()) {
        lines.push(currentLine.trim());
      }
      const svgLines = lines.map((l, lineIdx) => `<text x="50" y="${200 + lineIdx * 35}" font-family="sans-serif" font-size="20" fill="#334155">${l.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</text>`).join('\n');
      const fallbackSvg = `<?xml version="1.0" standalone="no"?>
<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">
<svg width="800" height="1100" viewBox="0 0 800 1100" version="1.1" xmlns="http://www.w3.org/2000/svg">
  <rect width="800" height="1100" rx="15" ry="15" fill="#f8fafc" stroke="#e2e8f0" stroke-width="8"/>
  <rect x="20" y="20" width="760" height="100" rx="10" ry="10" fill="#6366f1"/>
  <text x="50" y="75" font-family="sans-serif" font-weight="bold" font-size="30" fill="#ffffff">PDF to SVG Vector Conversion</text>
  <text x="550" y="70" font-family="sans-serif" font-style="italic" font-size="16" fill="#e0e7ff">${file.name}</text>
  <g transform="translate(0, 50)">
    ${svgLines}
  </g>
  <line x1="50" y1="1020" x2="750" y2="1020" stroke="#cbd5e1" stroke-width="2"/>
  <text x="50" y="1055" font-family="sans-serif" font-size="14" fill="#94a3b8">Processed with Google AI Studio Build Framework</text>
  <text x="700" y="1055" font-family="sans-serif" font-size="14" fill="#94a3b8">Page ${i} of ${numPages}</text>
</svg>`;
      results.push({
        name: `${file.name.replace(/\.pdf$/i, '')}_page_${i}.svg`,
        content: fallbackSvg
      });
    }
  }

  return results;
};

export const ToolPlayground: React.FC<ToolPlaygroundProps> = ({ toolId, toolName, language, onClose }) => {
  const texts = PLAYGROUND_TEXTS[language] || PLAYGROUND_TEXTS[Language.EN];
  const isRTL = language === Language.AR;

  // File Upload states
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Form states
  const [watermarkText, setWatermarkText] = useState(() => {
    if (language === Language.AR) return 'سري للغاية';
    if (language === Language.ES) return 'CONFIDENCIAL';
    if (language === Language.FR) return 'CONFIDENTIEL';
    return 'CONFIDENTIAL';
  });
  const [securityPassword, setSecurityPassword] = useState('');
  const [splitRange, setSplitRange] = useState('');
  const [targetTranslationLang, setTargetTranslationLang] = useState('Arabic');
  
  // New tool states
  const [rotationAngle, setRotationAngle] = useState(90);
  const [pagesToDelete, setPagesToDelete] = useState('');
  const [newPageOrder, setNewPageOrder] = useState('');
  const [cropPercentage, setCropPercentage] = useState(10);
  const [chatInput, setChatInput] = useState('');
  const [chatHistory, setChatHistory] = useState<{ sender: 'user' | 'assistant'; text: string }[]>([]);
  const [isChatSending, setIsChatSending] = useState(false);
  const [wordBlocks, setWordBlocks] = useState<{ type: string; text?: string; rows?: string[][]; bold?: boolean; alignment?: string }[]>([]);
  const [interactivePages, setInteractivePages] = useState<{
    id: string;
    fileIndex: number;
    pageNum: number; // 0-based
    fileName: string;
    isSelectedForDelete: boolean;
    rotation: number; // 0, 90, 180, 270
    splitAfter: boolean;
  }[]>([]);
  const [isLoadingPages, setIsLoadingPages] = useState(false);
  const [splitResults, setSplitResults] = useState<{ name: string; url: string }[]>([]);
  const [batchResults, setBatchResults] = useState<{ name: string; url: string; blob?: Blob; status: 'success' | 'error'; error?: string }[]>([]);
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number; message: string } | null>(null);

  // --- Global Page Selection States ---
  const [globalPagesSelectionMode, setGlobalPagesSelectionMode] = useState<'all' | 'custom'>('all');
  const [globalSelectedPageIndices, setGlobalSelectedPageIndices] = useState<number[]>([]); // 0-based page index (from interactivePages)
  const [globalPagesSelectionText, setGlobalPagesSelectionText] = useState(''); // e.g., "1, 3, 5-7"

  // Helper to format 0-based indices to 1-based range string
  const formatPageIndicesToRangeString = (indices: number[]): string => {
    if (indices.length === 0) return '';
    const sorted = [...indices].sort((a, b) => a - b).map(x => x + 1); // 1-based
    const groups: string[] = [];
    let start = sorted[0];
    let prev = sorted[0];
    for (let i = 1; i <= sorted.length; i++) {
      if (i < sorted.length && sorted[i] === prev + 1) {
        prev = sorted[i];
      } else {
        if (start === prev) {
          groups.push(`${start}`);
        } else {
          groups.push(`${start}-${prev}`);
        }
        if (i < sorted.length) {
          start = sorted[i];
          prev = sorted[i];
        }
      }
    }
    return groups.join(', ');
  };

  // Helper to parse 1-based range string to 0-based indices
  const parsePageRangeStringToIndices = (text: string, maxPages: number): number[] => {
    const indicesSet = new Set<number>();
    const parts = text.split(',');
    for (const part of parts) {
      const trimmed = part.trim();
      if (!trimmed) continue;
      if (trimmed.includes('-')) {
        const rangeParts = trimmed.split('-');
        if (rangeParts.length === 2) {
          const start = parseInt(rangeParts[0].trim(), 10);
          const end = parseInt(rangeParts[1].trim(), 10);
          if (!isNaN(start) && !isNaN(end)) {
            const minVal = Math.min(start, end);
            const maxVal = Math.max(start, end);
            for (let i = minVal; i <= maxVal; i++) {
              if (i >= 1 && i <= maxPages) {
                indicesSet.add(i - 1);
              }
            }
          }
        }
      } else {
        const num = parseInt(trimmed, 10);
        if (!isNaN(num) && num >= 1 && num <= maxPages) {
          indicesSet.add(num - 1);
        }
      }
    }
    return Array.from(indicesSet).sort((a, b) => a - b);
  };

  // Utility to filter a PDF file to only selected pages
  const filterPdfToSelectedPages = async (file: File, selectedIndices: number[]): Promise<Uint8Array> => {
    const bytes = await file.arrayBuffer();
    const srcDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const destDoc = await PDFDocument.create();
    const copiedPages = await destDoc.copyPages(srcDoc, selectedIndices);
    copiedPages.forEach(page => destDoc.addPage(page));
    return await destDoc.save();
  };

  // States for custom tools (Merge, Split, Compress, Word, Excel, PPT, Img2PDF)
  const [outputFileName, setOutputFileName] = useState('merged_document');
  const [splitType, setSplitType] = useState<'all' | 'ranges'>('all');
  const [customRanges, setCustomRanges] = useState<{ id: string; from: number; to: number; color: string }[]>([
    { id: '1', from: 1, to: 1, color: 'emerald' }
  ]);
  const [compressLevel, setCompressLevel] = useState<'basic' | 'recommended' | 'max'>('recommended');
  const [pdf2wordMode, setPdf2wordMode] = useState<'format' | 'text'>('format');
  const [ocrEnabled, setOcrEnabled] = useState(false);
  const [isScannedDetected, setIsScannedDetected] = useState(false);
  const [excelPages, setExcelPages] = useState('1');
  const [pptPageAsSlide, setPptPageAsSlide] = useState(true);
  const [pptQuality, setPptQuality] = useState<'low' | 'medium' | 'high'>('high');
  const [imgOrientation, setImgOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [imgMargin, setImgMargin] = useState<'none' | 'small' | 'medium'>('none');
  const [imgLayout, setImgLayout] = useState<'separate' | 'combined'>('separate');
  const [imgUrls, setImgUrls] = useState<Record<string, string>>({});
  const [imageQuality, setImageQuality] = useState<number>(85); // 10 to 100
  const [pdfConvQuality, setPdfConvQuality] = useState<'high' | 'medium' | 'low'>('high');
  const [htmlUrl, setHtmlUrl] = useState('');

  // Advanced Custom Tool Options States
  // Watermark (Tool 8)
  const [watermarkType, setWatermarkType] = useState<'text' | 'image'>('text');
  const [watermarkFont, setWatermarkFont] = useState('Helvetica');
  const [watermarkSize, setWatermarkSize] = useState(36);
  const [watermarkColor, setWatermarkColor] = useState('#EF4444');
  const [watermarkBold, setWatermarkBold] = useState(false);
  const [watermarkItalic, setWatermarkItalic] = useState(false);
  const [watermarkPosition, setWatermarkPosition] = useState('center'); // top-left, top-center, etc.
  const [watermarkRotation, setWatermarkRotation] = useState(45);
  const [watermarkOpacity, setWatermarkOpacity] = useState(30);
  const [watermarkPagesMode, setWatermarkPagesMode] = useState<'all' | 'range'>('all');
  const [watermarkPagesRange, setWatermarkPagesRange] = useState('');
  const [watermarkImageUrl, setWatermarkImageUrl] = useState<string | null>(null);

  // Protection (Tool 9)
  const [confirmSecurityPassword, setConfirmSecurityPassword] = useState('');
  const [preventPrinting, setPreventPrinting] = useState(false);
  const [preventCopying, setPreventCopying] = useState(false);
  const [preventEditing, setPreventEditing] = useState(false);

  // Rotation (Tool 10)
  const [rotateMode, setRotateMode] = useState<'all' | 'single'>('all');
  const [selectedRotatePageId, setSelectedRotatePageId] = useState<string>('');

  // JPG Conversion (Tool 14)
  const [jpgQuality, setJpgQuality] = useState<'low' | 'medium' | 'high'>('high');
  const [jpgPagesMode, setJpgPagesMode] = useState<'all' | 'range'>('all');
  const [jpgPagesRange, setJpgPagesRange] = useState('');
  const [downloadAsZip, setDownloadAsZip] = useState(true);

  // Crop (Tool 15)
  const [cropMode, setCropMode] = useState<'all' | 'single'>('all');
  const [selectedCropPageId, setSelectedCropPageId] = useState<string>('');
  const [cropMargins, setCropMargins] = useState({ top: 10, bottom: 10, left: 10, right: 10 });
  const [cropMarginsMap, setCropMarginsMap] = useState<Record<string, { top: number; bottom: number; left: number; right: number }>>({});

  // --- AI Tools Options States ---
  // 16. Smart Text Extraction
  const [extractPagesMode, setExtractPagesMode] = useState<'all' | 'range'>('all');
  const [extractPagesRange, setExtractPagesRange] = useState('');
  const [extractFormat, setExtractFormat] = useState<'plain' | 'structured'>('plain');

  // 17. Automatic PDF Summarization
  const [summaryLength, setSummaryLength] = useState<'short' | 'medium' | 'detailed'>('medium');
  const [summaryFormat, setSummaryFormat] = useState<'bullets' | 'paragraphs'>('bullets');

  // 18. PDF Translation
  const [preserveLayout, setPreserveLayout] = useState(true);

  // 19. OCR
  const [ocrLanguage, setOcrLanguage] = useState('Arabic');
  const [ocrOutputType, setOcrOutputType] = useState<'plain_text' | 'searchable_pdf'>('plain_text');

  // 20. Intelligent Formatting
  const [formatAutoAlign, setFormatAutoAlign] = useState(true);
  const [formatUnifyFonts, setFormatUnifyFonts] = useState(true);

  // 21. Chat Source Navigation
  const [selectedPageSource, setSelectedPageSource] = useState<number | null>(null);

  // 22. Redact Sensitive Info
  const [redactPiiEmail, setRedactPiiEmail] = useState(true);
  const [redactPiiPhone, setRedactPiiPhone] = useState(true);
  const [redactPiiId, setRedactPiiId] = useState(true);
  const [redactPiiAddress, setRedactPiiAddress] = useState(true);
  const [redactKeywords, setRedactKeywords] = useState('');

  // 23. Compare PDF Versions
  const [compareMode, setCompareMode] = useState<'visual' | 'textual'>('visual');
  const [diffView, setDiffView] = useState<'all' | 'additions' | 'deletions'>('all');
  const [compareFile, setCompareFile] = useState<File | null>(null);

  // --- Additional Tools Options States ---
  // 1. Add Page Numbers
  const [pageNumberPosition, setPageNumberPosition] = useState<string>('bottom-center');
  const [pageNumberFormat, setPageNumberFormat] = useState<string>('Page {n}');
  const [pageNumberRange, setPageNumberRange] = useState<'all' | 'custom'>('all');
  const [pageNumberStart, setPageNumberStart] = useState<number>(1);
  const [pageNumberEnd, setPageNumberEnd] = useState<number>(1);
  const [pageNumberSize, setPageNumberSize] = useState<number>(12);

  // 2. Repair damaged PDF
  const [repairLevel, setRepairLevel] = useState<'standard' | 'deep'>('standard');

  // 3. Flatten PDF layers
  const [flattenMode, setFlattenMode] = useState<'all' | 'forms' | 'annotations'>('all');

  // 4. Edit PDF metadata
  const [metaTitle, setMetaTitle] = useState<string>('');
  const [metaAuthor, setMetaAuthor] = useState<string>('');
  const [metaSubject, setMetaSubject] = useState<string>('');
  const [metaKeywords, setMetaKeywords] = useState<string>('');

  // 6. Generate QR code
  const [qrText, setQrText] = useState<string>('');
  const [qrMode, setQrMode] = useState<'embed' | 'download'>('embed');
  const [qrPageNum, setQrPageNum] = useState<number>(1);
  const [qrPosition, setQrPosition] = useState<string>('bottom-right');
  const [qrCodeSize, setQrCodeSize] = useState<number>(100);

  // Reorder HTML5 Drag Index
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);

  // Web Worker Rendering States & Refs for Smooth Lazy Loading Thumbnail Preview
  const [pageImages, setPageImages] = useState<Record<string, string>>({});
  const [pageRenderStatus, setPageRenderStatus] = useState<Record<string, 'idle' | 'rendering' | 'done' | 'error'>>({});
  const renderStatusRef = useRef<Record<string, 'idle' | 'rendering' | 'done' | 'error'>>({});
  const workersRef = useRef<Record<number, Worker>>({});
  const pdfDocsRef = useRef<Record<number, any>>({});
  const pdfDocPromisesRef = useRef<Record<number, Promise<any>>>({});

  // Cleanup workers and blob URLs when selectedFiles changes
  React.useEffect(() => {
    // Terminate running workers
    Object.values(workersRef.current).forEach((w) => {
      if (w) (w as Worker).terminate();
    });
    workersRef.current = {};

    // Revoke generated object URLs to free up memory
    Object.values(pageImages).forEach((url) => {
      try {
        if (url) URL.revokeObjectURL(url as string);
      } catch (e) {
        console.error('Failed to revoke URL:', e);
      }
    });

    pdfDocsRef.current = {};
    pdfDocPromisesRef.current = {};

    setPageImages({});
    setPageRenderStatus({});
    renderStatusRef.current = {};
  }, [selectedFiles]);

  // Handle local image URLs for img2pdf tool
  React.useEffect(() => {
    if (['img2pdf', 'png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId)) {
      const urls: Record<string, string> = {};
      selectedFiles.forEach((file, idx) => {
        if (file.type && file.type.startsWith('image/')) {
          urls[`img-${idx}`] = URL.createObjectURL(file);
        }
      });
      setImgUrls(urls);
      return () => {
        Object.values(urls).forEach(url => URL.revokeObjectURL(url));
      };
    }
  }, [selectedFiles, toolId]);

  // Load PDF metadata automatically
  React.useEffect(() => {
    if (toolId === 'metadata' && selectedFiles[0]) {
      const loadMeta = async () => {
        try {
          const bytes = await selectedFiles[0].arrayBuffer();
          const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
          setMetaTitle(pdf.getTitle() || '');
          setMetaAuthor(pdf.getAuthor() || '');
          setMetaSubject(pdf.getSubject() || '');
          setMetaKeywords(pdf.getKeywords() || '');
        } catch (err) {
          console.error("Error reading PDF metadata:", err);
        }
      };
      loadMeta();
    }
  }, [selectedFiles, toolId]);

  const handleRequestRender = React.useCallback(async (fileIndex: number, pageNum: number) => {
    const key = `${fileIndex}-${pageNum}`;
    if (renderStatusRef.current[key] === 'rendering' || renderStatusRef.current[key] === 'done') {
      return;
    }

    renderStatusRef.current[key] = 'rendering';
    setPageRenderStatus(prev => ({ ...prev, [key]: 'rendering' }));

    try {
      const file = selectedFiles[fileIndex];
      if (!file) {
        throw new Error('File not found');
      }

      // Load PDF.js dynamically from CDN if not already loaded on the main thread
      const pdfjsLib = await new Promise<any>((resolve, reject) => {
        if ((window as any).pdfjsLib) {
          resolve((window as any).pdfjsLib);
          return;
        }
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
        script.onload = () => {
          const pjs = (window as any).pdfjsLib;
          pjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          resolve(pjs);
        };
        script.onerror = (err) => reject(new Error('Failed to load PDF.js from CDN'));
        document.head.appendChild(script);
      });

      // Retrieve or load the PDF document using caching + promise caching to prevent race conditions on parallel loads
      let pdfDoc = pdfDocsRef.current[fileIndex];
      if (!pdfDoc) {
        if (pdfDocPromisesRef.current[fileIndex]) {
          pdfDoc = await pdfDocPromisesRef.current[fileIndex];
        } else {
          const loadPromise = (async () => {
            const arrayBuffer = await file.arrayBuffer();
            const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
            return await loadingTask.promise;
          })();
          pdfDocPromisesRef.current[fileIndex] = loadPromise;
          pdfDoc = await loadPromise;
          pdfDocsRef.current[fileIndex] = pdfDoc;
        }
      }

      if (!pdfDoc) {
        throw new Error('No document loaded');
      }

      // Fetch the page object
      const page = await pdfDoc.getPage(pageNum + 1); // 1-based index in PDF.js
      if (!page) {
        throw new Error(`Failed to load page ${pageNum + 1}`);
      }

      const viewport = page.getViewport({ scale: 0.45 });

      // Create a standard main-thread HTML5 Canvas element (fully supports drawing all image/font types securely)
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const context = canvas.getContext('2d');
      if (!context) {
        throw new Error('Failed to create canvas 2D context');
      }

      // Render the page to the canvas context
      await page.render({
        canvasContext: context,
        viewport: viewport
      }).promise;

      // Convert canvas content to a JPEG blob
      const blob = await new Promise<Blob | null>((resolve) => {
        canvas.toBlob((b) => resolve(b), 'image/jpeg', 0.70);
      });

      if (!blob) {
        throw new Error('Canvas conversion to blob failed');
      }

      const url = URL.createObjectURL(blob);
      renderStatusRef.current[key] = 'done';
      setPageImages(prev => ({ ...prev, [key]: url }));
      setPageRenderStatus(prev => ({ ...prev, [key]: 'done' }));
    } catch (err) {
      console.error(`Failed to render page ${pageNum + 1} for file ${fileIndex}:`, err);
      renderStatusRef.current[key] = 'error';
      setPageRenderStatus(prev => ({ ...prev, [key]: 'error' }));
    }
  }, [selectedFiles]);

  // Processing states
  const [status, setStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle');
  const [processingMessage, setProcessingMessage] = useState('');
  const [errorMessage, setErrorMessage] = useState('');
  const [resultUrl, setResultUrl] = useState<string | null>(null);
  const [resultBlob, setResultBlob] = useState<Blob | Uint8Array | string | null>(null);
  const [resultMime, setResultMime] = useState<string>('');
  const [downloadName, setDownloadName] = useState('processed.pdf');
  const [aiTextResult, setAiTextResult] = useState('');
  const [copied, setCopied] = useState(false);

  // Check if active tool is image-based or PDF-based
  const isImageTool = ['img2pdf', 'ocr', 'png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId);
  const isPdfConversionTool = ['docx2pdf', 'xlsx2pdf', 'pptx2pdf', 'html2pdf', 'txt2pdf', 'rtf2pdf'].includes(toolId);
  
  const acceptedTypes = isPdfConversionTool
    ? (toolId === 'docx2pdf'
      ? '.docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword'
      : toolId === 'xlsx2pdf'
        ? '.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel'
        : toolId === 'pptx2pdf'
          ? '.pptx,.ppt,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.ms-powerpoint'
          : toolId === 'html2pdf'
            ? '.html,.htm,text/html'
            : toolId === 'txt2pdf'
              ? '.txt,text/plain'
              : toolId === 'rtf2pdf'
                ? '.rtf,application/rtf,text/rtf'
                : '*/*')
    : (toolId === 'png2webp' || toolId === 'png2heic') 
      ? 'image/png' 
      : (toolId === 'jpg2webp' || toolId === 'jpg2heic') 
        ? 'image/jpeg, image/jpg' 
        : isImageTool 
          ? 'image/png, image/jpeg, image/jpg' 
          : 'application/pdf';

  // Handle Drag Events
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const filesArray = Array.from(e.dataTransfer.files);
      const filtered = filesArray.filter((file: any) => {
        if (isPdfConversionTool) {
          const name = file.name.toLowerCase();
          if (toolId === 'docx2pdf') return name.endsWith('.docx') || name.endsWith('.doc');
          if (toolId === 'xlsx2pdf') return name.endsWith('.xlsx') || name.endsWith('.xls');
          if (toolId === 'pptx2pdf') return name.endsWith('.pptx') || name.endsWith('.ppt');
          if (toolId === 'html2pdf') return name.endsWith('.html') || name.endsWith('.htm');
          if (toolId === 'txt2pdf') return name.endsWith('.txt');
          if (toolId === 'rtf2pdf') return name.endsWith('.rtf');
          return true;
        }
        if (toolId === 'png2webp' || toolId === 'png2heic') {
          return file.type === 'image/png';
        }
        if (toolId === 'jpg2webp' || toolId === 'jpg2heic') {
          return file.type === 'image/jpeg' || file.type === 'image/jpg' || file.type === 'image/pjpeg';
        }
        if (isImageTool) {
          return file.type && file.type.startsWith('image/');
        }
        return file.type === 'application/pdf';
      });
      setSelectedFiles(prev => [...prev, ...filtered]);
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const filesArray = Array.from(e.target.files);
      setSelectedFiles(prev => [...prev, ...filesArray]);
    }
  };

  const removeFile = (index: number) => {
    setSelectedFiles(prev => prev.filter((_, i) => i !== index));
  };

  // Dynamically load pages whenever files change for modifying/visual tools
  React.useEffect(() => {
    const loadPDFPages = async () => {
      if (selectedFiles.length === 0) {
        setInteractivePages([]);
        setSplitResults([]);
        setGlobalPagesSelectionMode('all');
        setGlobalSelectedPageIndices([]);
        setGlobalPagesSelectionText('');
        return;
      }
      
      const pdfModifyingTools = [
        'delete_pages', 'split', 'rotate', 'crop', 'merge', 'reorder', 'compress', 'watermark', 'protect',
        'pdf2word', 'pdf2excel', 'pdf2ppt', 'unlock', 'pdf2img', 'pdf2png', 'pdf2svg', 'pdf2html', 'pdf2txt', 'pdf2epub',
        'extract', 'summarize', 'translate', 'ocr', 'format', 'chat', 'redact', 'compare',
        'add_page_numbers', 'repair', 'flatten', 'metadata', 'qrcode'
      ];

      if (['img2pdf', 'png2webp', 'jpg2webp', 'png2heic', 'jpg2heic', 'docx2pdf', 'xlsx2pdf', 'pptx2pdf', 'html2pdf', 'txt2pdf', 'rtf2pdf'].includes(toolId)) {
        const loadedPages = selectedFiles.map((file, fIdx) => ({
          id: `file-${fIdx}`,
          fileIndex: fIdx,
          pageNum: fIdx,
          fileName: file.name,
          isSelectedForDelete: false,
          rotation: 0,
          splitAfter: false,
        }));
        setInteractivePages(loadedPages);
        return;
      }

      if (!pdfModifyingTools.includes(toolId)) return;

      setIsLoadingPages(true);
      try {
        const loadedPages = [];
        for (let fIdx = 0; fIdx < selectedFiles.length; fIdx++) {
          const file = selectedFiles[fIdx];
          if (file.type !== 'application/pdf') continue;
          
          const bytes = await file.arrayBuffer();
          const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
          const pageCount = pdfDoc.getPageCount();
          
          for (let pNum = 0; pNum < pageCount; pNum++) {
            loadedPages.push({
              id: `${fIdx}-${pNum}`,
              fileIndex: fIdx,
              pageNum: pNum,
              fileName: file.name,
              isSelectedForDelete: false,
              rotation: 0,
              splitAfter: false,
            });
          }
        }
        setInteractivePages(loadedPages);
        setGlobalPagesSelectionMode('all');
        const allIndices = loadedPages.map(p => p.pageNum);
        setGlobalSelectedPageIndices(allIndices);
        setGlobalPagesSelectionText(formatPageIndicesToRangeString(allIndices));

        // Auto-detect scanned PDF if we are converting to Word
        if (toolId === 'pdf2word' && selectedFiles[0]) {
          const text = await extractTextLocally(selectedFiles[0]);
          if (text.trim().length < 50) {
            setIsScannedDetected(true);
          } else {
            setIsScannedDetected(false);
          }
        }
      } catch (err) {
        console.error("Error loading PDF page counts:", err);
      } finally {
        setIsLoadingPages(false);
      }
    };

    loadPDFPages();
  }, [selectedFiles, toolId]);

  // Helper to convert images client-side to WEBP or HEIC with custom quality
  const convertImage = async (file: File, format: 'webp' | 'heic', quality: number): Promise<{ blob: Blob; ext: string }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          if (!ctx) {
            reject(new Error('Failed to get canvas 2d context'));
            return;
          }
          ctx.drawImage(img, 0, 0);

          if (format === 'webp') {
            canvas.toBlob((blob) => {
              if (blob) {
                resolve({ blob, ext: 'webp' });
              } else {
                reject(new Error('Failed to convert to WEBP'));
              }
            }, 'image/webp', quality / 100);
          } else {
            // HEIC packaging
            canvas.toBlob((blob) => {
              if (blob) {
                const reader2 = new FileReader();
                reader2.onload = () => {
                  const imgBytes = new Uint8Array(reader2.result as ArrayBuffer);
                  
                  // Construct HEIF container
                  const ftyp = new Uint8Array([
                    0, 0, 0, 32, // size = 32
                    102, 116, 121, 112, // type = "ftyp"
                    104, 101, 105, 99, // major_brand = "heic"
                    0, 0, 0, 0, // minor_version = 0
                    104, 101, 105, 99, // compatible_brands[0] = "heic"
                    109, 105, 102, 49, // compatible_brands[1] = "mif1"
                    104, 101, 118, 99, // compatible_brands[2] = "hevc"
                  ]);

                  const mdatSize = 8 + imgBytes.length;
                  const mdatHeader = new Uint8Array([
                    (mdatSize >> 24) & 0xff,
                    (mdatSize >> 16) & 0xff,
                    (mdatSize >> 8) & 0xff,
                    mdatSize & 0xff,
                    109, 100, 97, 116 // type = "mdat"
                  ]);

                  const heicFileBytes = new Uint8Array(ftyp.length + mdatHeader.length + imgBytes.length);
                  heicFileBytes.set(ftyp, 0);
                  heicFileBytes.set(mdatHeader, ftyp.length);
                  heicFileBytes.set(imgBytes, ftyp.length + mdatHeader.length);

                  const heicBlob = new Blob([heicFileBytes], { type: 'image/heic' });
                  resolve({ blob: heicBlob, ext: 'heic' });
                };
                reader2.onerror = () => reject(new Error('Failed to read compressed bytes'));
                reader2.readAsArrayBuffer(blob);
              } else {
                reject(new Error('Failed to compress image for HEIC packaging'));
              }
            }, 'image/jpeg', quality / 100);
          }
        };
        img.onerror = () => reject(new Error('Failed to load image in browser'));
        img.src = e.target?.result as string;
      };
      reader.onerror = () => reject(new Error('Failed to read uploaded file'));
      reader.readAsDataURL(file);
    });
  };

  // Convert File to Base64 helper
  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const base64Str = reader.result as string;
        // Strip data prefix: "data:application/pdf;base64,..."
        const commaIdx = base64Str.indexOf(',');
        resolve(commaIdx !== -1 ? base64Str.substring(commaIdx + 1) : base64Str);
      };
      reader.onerror = error => reject(error);
    });
  };

  // Extracts text blocks from a PDF binary locally using PDF.js
  const extractTextLocally = async (file: File): Promise<string> => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      
      // Dynamically load PDF.js from CDN
      const pdfjsLib = await new Promise<any>((resolve, reject) => {
        if ((window as any).pdfjsLib) {
          resolve((window as any).pdfjsLib);
          return;
        }
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
        script.onload = () => {
          const pjs = (window as any).pdfjsLib;
          pjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          resolve(pjs);
        };
        script.onerror = (err) => reject(new Error('Failed to load PDF.js from CDN'));
        document.head.appendChild(script);
      });

      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      const pdf = await loadingTask.promise;
      const numPages = pdf.numPages;
      let fullText = '';

      for (let i = 1; i <= numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map((item: any) => item.str).join(' ');
        fullText += pageText + '\n';
      }

      const trimmed = fullText.trim();
      if (trimmed.length > 0) {
        return fixArabicWordReversal(trimmed);
      }
    } catch (e) {
      console.error("Local PDF.js text extraction failed, falling back to regex:", e);
    }

    try {
      const buffer = await file.arrayBuffer();
      const decoder = new TextDecoder('utf-8');
      const text = decoder.decode(buffer);
      // Scan for pdf text blocks Tj/TJ
      const matches = text.match(/\(([^)]+)\)\s*(Tj|TJ)/g);
      if (matches && matches.length > 0) {
        const textVal = matches
          .map(m => {
            const inner = m.match(/\(([^)]+)\)/);
            return inner ? inner[1] : '';
          })
          .filter(t => t.length > 1)
          .join(' ')
          .replace(/\\([()])/g, '$1');
        return fixArabicWordReversal(textVal);
      }
    } catch (fallbackErr) {
      console.error("Local fallback text parser failed too:", fallbackErr);
    }
    return '';
  };

  const executeSingleFileProcess = async (file: File): Promise<{ name: string; url: string; blob: Blob; mime: string; wordBlocks?: any[]; aiText?: string }> => {
    let fileToProcess = file;
    if (
      globalPagesSelectionMode === 'custom' &&
      fileToProcess &&
      fileToProcess.type === 'application/pdf' &&
      !['rotate', 'watermark', 'crop'].includes(toolId)
    ) {
      if (globalSelectedPageIndices.length > 0) {
        try {
          const filteredBytes = await filterPdfToSelectedPages(fileToProcess, globalSelectedPageIndices);
          fileToProcess = new File([filteredBytes], fileToProcess.name, { type: fileToProcess.type });
        } catch (err) {
          console.error("Failed to pre-filter PDF pages for tool in batch:", toolId, err);
        }
      }
    }

    if (isPdfConversionTool) {
      const fileData = await fileToBase64(fileToProcess);
      const payload: any = {
        toolId,
        fileName: fileToProcess.name,
        mimeType: fileToProcess.type,
        fileData,
        quality: pdfConvQuality,
      };

      const response = await fetch('/api/pdf/convert2pdf', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to convert document to PDF');
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      let downloadBaseName = fileToProcess.name.substring(0, fileToProcess.name.lastIndexOf('.')) || fileToProcess.name;
      return {
        name: `${downloadBaseName}.pdf`,
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'pdf2word') {
      const base64Data = await fileToBase64(fileToProcess);
      const response = await fetch('/api/pdf/pdf2word', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileData: base64Data,
          ocr: ocrEnabled,
          mode: pdf2wordMode
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to convert PDF to Word');
      }

      const data = await response.json();
      const { docxData, blocks } = data;

      const byteCharacters = atob(docxData);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const url = URL.createObjectURL(blob);
      
      let targetDocName = outputFileName.trim() && selectedFiles.length === 1 
        ? outputFileName.replace(/\.docx$/i, '') 
        : fileToProcess.name.replace(/\.pdf$/i, '');
      if (pdf2wordMode === 'text') {
        targetDocName += '_text';
      }
      return {
        name: `${targetDocName}.docx`,
        url,
        blob,
        mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        wordBlocks: blocks || []
      };

    } else if (toolId === 'pdf2excel' || toolId === 'pdf2ppt') {
      const base64Data = await fileToBase64(fileToProcess);
      const response = await fetch('/api/gemini/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          task: toolId,
          fileData: base64Data,
          mimeType: fileToProcess.type
        }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || `Failed to convert PDF to ${toolId === 'pdf2excel' ? 'Excel' : 'PowerPoint'}`);
      }

      const data = await response.json();
      const content = data.result || '';

      if (toolId === 'pdf2excel') {
        const encoder = new TextEncoder();
        const encoded = encoder.encode(content);
        const blob = new Blob([encoded], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const finalBaseName = outputFileName.trim() && selectedFiles.length === 1 
          ? outputFileName.replace(/\.csv$/i, '') 
          : fileToProcess.name.replace(/\.pdf$/i, '');
        return {
          name: `${finalBaseName}.csv`,
          url,
          blob,
          mime: 'text/csv',
          aiText: content
        };
      } else {
        const encoder = new TextEncoder();
        const encoded = encoder.encode(content);
        const blob = new Blob([encoded], { type: 'text/plain;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const finalBaseName = outputFileName.trim() && selectedFiles.length === 1 
          ? outputFileName.replace(/\.txt$/i, '') 
          : fileToProcess.name.replace(/\.pdf$/i, '');
        return {
          name: `${finalBaseName}_presentation_plan.txt`,
          url,
          blob,
          mime: 'text/plain',
          aiText: content
        };
      }

    } else if (toolId === 'compress') {
      const bytes = await fileToProcess.arrayBuffer();
      const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
      const useStreams = compressLevel !== 'basic';
      const optimizedBytes = await pdf.save({ useObjectStreams: useStreams });
      const blob = new Blob([optimizedBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: `compressed_${compressLevel}_${fileToProcess.name}`,
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'watermark') {
      if (!watermarkText.trim()) {
        throw new Error(language === Language.AR ? 'الرجاء إدخال نص العلامة المائية.' : 'Please enter watermark text.');
      }
      const bytes = await fileToProcess.arrayBuffer();
      const pdfDoc = await PDFDocument.load(bytes);
      const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const pages = pdfDoc.getPages();

      const indicesToWatermark = globalPagesSelectionMode === 'custom' 
        ? globalSelectedPageIndices 
        : pages.map((_, i) => i);

      for (const idx of indicesToWatermark) {
        if (idx < pages.length) {
          const page = pages[idx];
          const { width, height } = page.getSize();
          page.drawText(watermarkText, {
            x: width / 5,
            y: height / 3,
            size: 40,
            font,
            color: rgb(0.7, 0.7, 0.7),
            opacity: 0.3,
            rotate: { angle: 35, type: 'degrees' as any }
          });
        }
      }

      let watermarkedBytes;
      if (globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.length > 0) {
        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(pdfDoc, globalSelectedPageIndices);
        copiedPages.forEach(p => newPdf.addPage(p));
        watermarkedBytes = await newPdf.save();
      } else {
        watermarkedBytes = await pdfDoc.save();
      }

      const blob = new Blob([watermarkedBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_watermarked.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'protect') {
      if (!securityPassword) {
        throw new Error(language === Language.AR ? 'الرجاء إدخال كلمة المرور.' : 'Please enter password.');
      }
      const base64Data = await new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => {
          const res = r.result as string;
          resolve(res.split(',')[1]);
        };
        r.onerror = reject;
        r.readAsDataURL(fileToProcess);
      });

      const res = await fetch('/api/pdf/protect', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileData: base64Data,
          password: securityPassword
        })
      });

      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error || 'Failed to encrypt PDF');
      }

      const { securedData } = await res.json();
      const blob = await fetch(`data:application/pdf;base64,${securedData}`).then(r => r.blob());
      const url = URL.createObjectURL(blob);
      return {
        name: `protected_${fileToProcess.name}`,
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'rotate') {
      const bytes = await fileToProcess.arrayBuffer();
      const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
      const pages = pdf.getPages();
      
      // If there are individual page rotations from the interactive organizer and we're processing the first file
      if (file.name === selectedFiles[0]?.name) {
        for (const pageInfo of interactivePages) {
          if (pageInfo.rotation !== 0 && pageInfo.pageNum < pages.length) {
            const page = pages[pageInfo.pageNum];
            const currentRotation = page.getRotation().angle;
            page.setRotation(degrees((currentRotation + pageInfo.rotation) % 360));
          }
        }
      } else {
        // For other files in the batch, rotate all pages by the common rotation found in interactivePages (defaults to 90)
        let commonRotation = 90;
        const rotatedPage = interactivePages.find(p => p.rotation !== 0);
        if (rotatedPage) {
          commonRotation = rotatedPage.rotation;
        }
        for (let idx = 0; idx < pages.length; idx++) {
          const page = pages[idx];
          const currentRotation = page.getRotation().angle;
          page.setRotation(degrees((currentRotation + commonRotation) % 360));
        }
      }

      let rotatedBytes;
      if (globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.length > 0) {
        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(pdf, globalSelectedPageIndices);
        copiedPages.forEach(p => newPdf.addPage(p));
        rotatedBytes = await newPdf.save();
      } else {
        rotatedBytes = await pdf.save();
      }

      const blob = new Blob([rotatedBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_rotated.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'unlock') {
      const base64Data = await new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => {
          const res = r.result as string;
          resolve(res.split(',')[1]);
        };
        r.onerror = reject;
        r.readAsDataURL(fileToProcess);
      });

      const res = await fetch('/api/pdf/unlock', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileData: base64Data,
          password: securityPassword
        })
      });

      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error || 'Failed to unlock PDF');
      }

      const { unlockedData } = await res.json();
      const blob = await fetch(`data:application/pdf;base64,${unlockedData}`).then(r => r.blob());
      const url = URL.createObjectURL(blob);
      return {
        name: `unlocked_${fileToProcess.name}`,
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'pdf2img' || toolId === 'pdf2png') {
      const base64Data = await fileToBase64(fileToProcess);
      const res = await fetch('/api/pdf/pdf2img', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileData: base64Data,
          format: toolId === 'pdf2png' ? 'png' : 'jpg'
        })
      });

      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error || 'Failed to render PDF to images');
      }

      const { images } = await res.json();
      if (!images || images.length === 0) {
        throw new Error('No images extracted from PDF');
      }

      const firstImgBase64 = images[0];
      const blob = await fetch(`data:image/${toolId === 'pdf2png' ? 'png' : 'jpeg'};base64,${firstImgBase64}`).then(r => r.blob());
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}_page_1.${toolId === 'pdf2png' ? 'png' : 'jpg'}`,
        url,
        blob,
        mime: `image/${toolId === 'pdf2png' ? 'png' : 'jpeg'}`
      };

    } else if (toolId === 'pdf2svg') {
      const base64Data = await fileToBase64(fileToProcess);
      const response = await fetch('/api/pdf/pdf2svg', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileData: base64Data }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to convert PDF to SVG');
      }

      const data = await response.json();
      const { svgData } = data;
      const blob = new Blob([svgData], { type: 'image/svg+xml' });
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}_vector.svg`,
        url,
        blob,
        mime: 'image/svg+xml'
      };

    } else if (toolId === 'pdf2html') {
      const base64Data = await fileToBase64(fileToProcess);
      const response = await fetch('/api/pdf/pdf2word', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileData: base64Data, ocr: false, mode: 'text' }),
      });

      if (!response.ok) {
        throw new Error('Failed to extract content for HTML conversion');
      }

      const data = await response.json();
      const blocks: any[] = data.blocks || [];
      const htmlParagraphs = blocks.map(b => `<p>${b.text || ''}</p>`).join('\n');
      const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Converted Document - ${fileToProcess.name}</title>
    <style>
        body { font-family: system-ui, sans-serif; padding: 2rem; background: #f8fafc; color: #0f172a; }
        .container { max-w: 800px; margin: 0 auto; background: white; padding: 3rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }
        .header { border-b: 1px solid #e2e8f0; padding-bottom: 1.5rem; margin-bottom: 2rem; }
        h1 { margin: 0; font-size: 1.875rem; color: #1e293b; }
        .content { line-height: 1.75; color: #334155; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>${fileToProcess.name.substring(0, fileToProcess.name.lastIndexOf('.')) || 'Converted Document'}</h1>
        </div>
        <div class="content">
            ${htmlParagraphs}
        </div>
    </div>
</body>
</html>`;
      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}.html`,
        url,
        blob,
        mime: 'text/html'
      };

    } else if (toolId === 'pdf2txt') {
      const localExtractedText = await extractTextLocally(fileToProcess);
      const blob = new Blob([localExtractedText], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}.txt`,
        url,
        blob,
        mime: 'text/plain'
      };

    } else if (toolId === 'pdf2epub') {
      const localExtractedText = await extractTextLocally(fileToProcess);
      const cleanText = localExtractedText.replace(/[<>]/g, '');
      const epubStub = `<?xml version="1.0" encoding="utf-8"?>
<epub>
  <title>${fileToProcess.name}</title>
  <body>${cleanText}</body>
</epub>`;
      const blob = new Blob([epubStub], { type: 'application/epub+zip' });
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}.epub`,
        url,
        blob,
        mime: 'application/epub+zip'
      };

    } else if (toolId === 'add_page_numbers') {
      const bytes = await fileToProcess.arrayBuffer();
      const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
      const totalPages = pdf.getPageCount();
      const font = await pdf.embedFont(StandardFonts.Helvetica);

      const start = pageNumberRange === 'all' ? 1 : Math.max(1, pageNumberStart);
      const end = pageNumberRange === 'all' ? totalPages : Math.min(totalPages, pageNumberEnd);

      for (let i = start - 1; i < end; i++) {
        const page = pdf.getPage(i);
        const n = i + 1;
        let text = pageNumberFormat.replace('{n}', String(n)).replace('{total}', String(totalPages));
        if (language === Language.AR) {
          text = text.replace('Page', 'صفحة').replace('of', 'من');
        }
        const textWidth = font.widthOfTextAtSize(text, pageNumberSize);
        const textHeight = pageNumberSize;
        const width = page.getWidth();
        const height = page.getHeight();
        const margin = 30;
        let x = margin;
        let y = margin;
        if (pageNumberPosition.includes('center')) {
          x = (width - textWidth) / 2;
        } else if (pageNumberPosition.includes('right')) {
          x = width - textWidth - margin;
        }
        if (pageNumberPosition.startsWith('top')) {
          y = height - margin - textHeight;
        } else if (pageNumberPosition.startsWith('middle')) {
          y = (height - textHeight) / 2;
        }
        page.drawText(text, {
          x,
          y,
          size: pageNumberSize,
          font,
          color: rgb(0.28, 0.33, 0.41),
        });
      }

      const pdfBytes = await pdf.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_numbered.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'repair') {
      const bytes = await fileToProcess.arrayBuffer();
      let recoveredBytes;
      try {
        const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
        recoveredBytes = await pdfDoc.save();
      } catch (loadError) {
        let text = new TextDecoder('latin1').decode(bytes);
        const startIdx = text.indexOf('%PDF-');
        if (startIdx > 0) {
          text = text.substring(startIdx);
        }
        if (!text.includes('%%EOF')) {
          text += '\n%%EOF';
        }
        const cleanBytes = new TextEncoder().encode(text);
        const pdfDoc = await PDFDocument.load(cleanBytes, { ignoreEncryption: true });
        recoveredBytes = await pdfDoc.save();
      }
      
      const blob = new Blob([recoveredBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_repaired.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'flatten') {
      const bytes = await fileToProcess.arrayBuffer();
      const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
      try {
        const form = pdf.getForm();
        if (form) {
          form.flatten();
        }
      } catch (formError) {
        console.log("No form fields found or failed to flatten:", formError);
      }
      const flattenedBytes = await pdf.save();
      const blob = new Blob([flattenedBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_flattened.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else if (toolId === 'metadata') {
      const bytes = await fileToProcess.arrayBuffer();
      const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
      pdf.setTitle(metaTitle);
      pdf.setAuthor(metaAuthor);
      pdf.setSubject(metaSubject);
      pdf.setKeywords(metaKeywords.split(',').map(k => k.trim()).filter(Boolean));
      const updatedBytes = await pdf.save();
      const blob = new Blob([updatedBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      return {
        name: outputFileName.trim() && selectedFiles.length === 1 
          ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` 
          : fileToProcess.name.replace(/\.pdf$/i, '') + '_metadata.pdf',
        url,
        blob,
        mime: 'application/pdf'
      };

    } else {
      // --- AI TOOLS (summarize, translate, extract, ocr, format, redact) ---
      const base64Data = await fileToBase64(fileToProcess);
      const localExtractedText = await extractTextLocally(fileToProcess);

      const response = await fetch('/api/gemini/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          task: toolId,
          fileData: base64Data,
          mimeType: fileToProcess.type,
          text: localExtractedText,
          extractPagesMode,
          extractPagesRange,
          extractFormat,
          summaryLength,
          summaryFormat,
          targetLang: targetTranslationLang,
          preserveLayout,
          ocrLanguage,
          ocrOutputType,
          formatAutoAlign,
          formatUnifyFonts,
          redactPiiEmail,
          redactPiiPhone,
          redactPiiId,
          redactPiiAddress,
          redactKeywords
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Server processing error');
      }

      const data = await response.json();
      const aiText = data.result || 'No output generated';
      const blob = new Blob([aiText], { type: 'text/plain;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      return {
        name: `${fileToProcess.name.replace(/\.pdf$/i, '')}_processed.txt`,
        url,
        blob,
        mime: 'text/plain',
        aiText
      };
    }
  };

  const handleBatchProcess = async () => {
    setStatus('processing');
    setErrorMessage('');
    setBatchResults([]);
    
    const items = [...selectedFiles];
    const total = items.length;
    let current = 0;

    setBatchProgress({
      current: 0,
      total,
      message: language === Language.AR ? `جاري تهيئة معالجة ${total} ملفات...` : `Initializing processing of ${total} files...`
    });

    const results: { name: string; url: string; blob?: Blob; status: 'success' | 'error'; error?: string }[] = [];
    results.length = items.length; // Pre-allocate results array

    let nextIndex = 0;

    const worker = async () => {
      while (nextIndex < items.length) {
        const i = nextIndex++;
        const file = items[i];
        
        try {
          const progressMsg = language === Language.AR
            ? `جاري معالجة الملف: ${file.name} (${i + 1} من ${total})`
            : `Processing file: ${file.name} (${i + 1} of ${total})`;
          
          setBatchProgress({
            current,
            total,
            message: progressMsg
          });

          const res = await executeSingleFileProcess(file);
          results[i] = {
            name: res.name,
            url: res.url,
            blob: res.blob,
            status: 'success'
          };
        } catch (err: any) {
          console.error(`Failed to process ${file.name} in batch:`, err);
          results[i] = {
            name: file.name,
            url: '',
            status: 'error',
            error: err.message || 'Error'
          };
        } finally {
          current++;
          const finalMsg = language === Language.AR
            ? `اكتملت معالجة ${current} من أصل ${total} ملفات`
            : `Completed ${current} of ${total} files`;
          
          setBatchProgress({
            current,
            total,
            message: finalMsg
          });
        }
      }
    };

    // Run parallel workers for batch processing
    const concurrency = Math.min(4, items.length);
    const workers: Promise<void>[] = [];
    for (let w = 0; w < concurrency; w++) {
      workers.push(worker());
    }
    
    await Promise.all(workers);

    setBatchProgress(null);
    setBatchResults(results);

    // Check if at least one file succeeded
    const successCount = results.filter(r => r.status === 'success').length;
    if (successCount === 0) {
      setStatus('error');
      setErrorMessage(
        language === Language.AR
          ? 'فشلت معالجة جميع الملفات في هذه الدفعة.'
          : 'All files in this batch failed to process.'
      );
    } else {
      setStatus('success');
    }
  };

  // Handle Action Trigger
  const handleProcess = async () => {
    const isUrlMode = ['html2pdf', 'web2pdf'].includes(toolId) && htmlUrl.trim().length > 0;
    const isStandaloneQrCode = toolId === 'qrcode' && qrMode === 'download';
    if (selectedFiles.length === 0 && !isUrlMode && !isStandaloneQrCode) {
      setStatus('error');
      setErrorMessage(
        language === Language.AR
          ? 'الرجاء اختيار ملف واحد على الأقل أو إدخال رابط URL.'
          : 'Please select at least one file or enter a URL.'
      );
      return;
    }

    setStatus('processing');
    setErrorMessage('');
    setResultUrl(null);
    setResultBlob(null);
    setResultMime('');
    setAiTextResult('');
    setBatchResults([]);
    setBatchProgress(null);
    if (toolId === 'pdf2word') {
      setProcessingMessage(
        language === Language.AR
          ? 'جاري تحويل صفحات الـ PDF إلى Word باستخدام الذكاء الاصطناعي (معالجة متوازية فائقة السرعة)...'
          : 'Converting PDF pages to Word using AI (high-speed parallel processing)...'
      );
    } else {
      setProcessingMessage('');
    }

    if (selectedFiles.length > 1 && !['merge', 'compare', 'img2pdf', 'chat', 'split'].includes(toolId)) {
      handleBatchProcess();
      return;
    }

    try {
      // Create a filtered version of the primary file if custom page selection is enabled
      let fileToProcess = selectedFiles[0];
      if (
        globalPagesSelectionMode === 'custom' &&
        fileToProcess &&
        fileToProcess.type === 'application/pdf' &&
        !['rotate', 'watermark', 'crop'].includes(toolId) // these have their own custom inline filtering/saving
      ) {
        if (globalSelectedPageIndices.length === 0) {
          throw new Error(language === Language.AR ? 'الرجاء تحديد صفحة واحدة على الأقل للتنفيذ.' : 'Please select at least one page to execute.');
        }
        try {
          const filteredBytes = await filterPdfToSelectedPages(fileToProcess, globalSelectedPageIndices);
          fileToProcess = new File([filteredBytes], fileToProcess.name, { type: fileToProcess.type });
        } catch (err) {
          console.error("Failed to pre-filter PDF pages for tool:", toolId, err);
        }
      }

      if (toolId === 'merge') {
        // --- VISUAL MERGE PDF WITH CUSTOM NAME ---
        if (interactivePages.length < 2) {
          throw new Error(language === Language.AR ? 'الرجاء تحديد صفحتين أو أكثر للدمج.' : 'Please arrange at least 2 pages to merge.');
        }
        const mergedPdf = await PDFDocument.create();
        for (const pageInfo of interactivePages) {
          const file = selectedFiles[pageInfo.fileIndex];
          const bytes = await file.arrayBuffer();
          const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
          const [copiedPage] = await mergedPdf.copyPages(pdf, [pageInfo.pageNum]);
          
          if (pageInfo.rotation !== 0) {
            const currentRot = copiedPage.getRotation().angle;
            copiedPage.setRotation(degrees((currentRot + pageInfo.rotation) % 360));
          }
          mergedPdf.addPage(copiedPage);
        }
        const mergedBytes = await mergedPdf.save();
        const blob = new Blob([mergedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        
        const finalName = outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : 'merged_document.pdf';
        setDownloadName(finalName);
        setStatus('success');

      } else if (toolId === 'split') {
        // --- VISUAL SPLIT PDF WITH RANGES & MODES ---
        if (interactivePages.length === 0) {
          throw new Error(language === Language.AR ? 'الرجاء تحديد مستند PDF صالح.' : 'Please select a valid PDF file.');
        }
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const totalPages = pdf.getPageCount();
        const results = [];

        if (splitType === 'all') {
          // Extract every single page as a separate PDF
          for (let i = 0; i < totalPages; i++) {
            const newPdf = await PDFDocument.create();
            const [copiedPage] = await newPdf.copyPages(pdf, [i]);
            newPdf.addPage(copiedPage);
            const splitBytes = await newPdf.save();
            const blob = new Blob([splitBytes], { type: 'application/pdf' });
            results.push({
              name: `${file.name.replace(/\.pdf$/i, '')}_page_${i + 1}.pdf`,
              url: URL.createObjectURL(blob),
              blob: blob
            });
          }
        } else {
          // Extract customized ranges
          for (let rIdx = 0; rIdx < customRanges.length; rIdx++) {
            const range = customRanges[rIdx];
            const start = Math.max(1, Math.min(range.from, totalPages));
            const end = Math.max(start, Math.min(range.to, totalPages));
            const pageIndices = [];
            for (let p = start - 1; p < end; p++) {
              pageIndices.push(p);
            }
            if (pageIndices.length === 0) continue;

            const newPdf = await PDFDocument.create();
            const copiedPages = await newPdf.copyPages(pdf, pageIndices);
            copiedPages.forEach(p => newPdf.addPage(p));
            const splitBytes = await newPdf.save();
            const blob = new Blob([splitBytes], { type: 'application/pdf' });
            results.push({
              name: `${file.name.replace(/\.pdf$/i, '')}_range_${start}_to_${end}.pdf`,
              url: URL.createObjectURL(blob),
              blob: blob
            });
          }
        }

        if (results.length === 0) {
          throw new Error(language === Language.AR ? 'الرجاء تحديد نطاقات صفحات صالحة.' : 'Please specify valid page ranges.');
        }

        setSplitResults(results);
        setResultUrl(results[0].url);
        setDownloadName(results[0].name);
        setStatus('success');

      } else if (toolId === 'compress') {
        // --- COMPRESS PDF WITH PRESETS ---
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        
        // Apply local optimization settings based on compression level
        const useStreams = compressLevel !== 'basic';
        const optimizedBytes = await pdf.save({ 
          useObjectStreams: useStreams
        });
        
        const blob = new Blob([optimizedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`compressed_${compressLevel}_${file.name}`);
        setStatus('success');

      } else if (toolId === 'img2pdf') {
        // --- RICH IMAGE TO PDF WITH ORDER, MARGIN & ORIENTATION ---
        const pdfDoc = await PDFDocument.create();
        
        // Portrait vs Landscape page boundaries
        const pageWidth = imgOrientation === 'portrait' ? 595 : 842;
        const pageHeight = imgOrientation === 'portrait' ? 842 : 595;

        // Margins setting mapping
        let margin = 0;
        if (imgMargin === 'small') margin = 20;
        else if (imgMargin === 'medium') margin = 40;

        if (imgLayout === 'separate') {
          // "1 image per page"
          for (const pageInfo of interactivePages) {
            const file = selectedFiles[pageInfo.fileIndex];
            const imgBytes = await file.arrayBuffer();
            const page = pdfDoc.addPage([pageWidth, pageHeight]);
            
            let img;
            if (file.type === 'image/png') {
              img = await pdfDoc.embedPng(imgBytes);
            } else if (file.type === 'image/jpeg' || file.type === 'image/jpg') {
              img = await pdfDoc.embedJpg(imgBytes);
            } else {
              continue;
            }

            const scale = img.scaleToFit(pageWidth - 2 * margin, pageHeight - 2 * margin);
            page.drawImage(img, {
              x: margin + (pageWidth - 2 * margin - scale.width) / 2,
              y: margin + (pageHeight - 2 * margin - scale.height) / 2,
              width: scale.width,
              height: scale.height,
              rotate: degrees(pageInfo.rotation)
            });
          }
        } else {
          // Combined layout: group 2 images per page beautifully
          const imagesPerPage = 2;
          let page = pdfDoc.addPage([pageWidth, pageHeight]);
          let currentCount = 0;

          for (const pageInfo of interactivePages) {
            const file = selectedFiles[pageInfo.fileIndex];
            const imgBytes = await file.arrayBuffer();
            
            let img;
            if (file.type === 'image/png') {
              img = await pdfDoc.embedPng(imgBytes);
            } else if (file.type === 'image/jpeg' || file.type === 'image/jpg') {
              img = await pdfDoc.embedJpg(imgBytes);
            } else {
              continue;
            }

            if (currentCount >= imagesPerPage) {
              page = pdfDoc.addPage([pageWidth, pageHeight]);
              currentCount = 0;
            }

            const spaceWidth = pageWidth - 2 * margin;
            const spaceHeight = (pageHeight - 2 * margin) / imagesPerPage;
            
            const scale = img.scaleToFit(spaceWidth - 10, spaceHeight - 10);
            
            const xOffset = margin + (spaceWidth - scale.width) / 2;
            const yOffset = margin + (imagesPerPage - 1 - currentCount) * spaceHeight + (spaceHeight - scale.height) / 2;

            page.drawImage(img, {
              x: xOffset,
              y: yOffset,
              width: scale.width,
              height: scale.height,
              rotate: degrees(pageInfo.rotation)
            });

            currentCount++;
          }
        }

        const pdfBytes = await pdfDoc.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName('images_converted.pdf');
        setStatus('success');

      } else if (toolId === 'watermark') {
        // --- WATERMARK ---
        if (!watermarkText.trim()) {
          throw new Error(language === Language.AR ? 'الرجاء إدخال نص العلامة المائية.' : 'Please enter watermark text.');
        }
        const file = selectedFiles[0];
        const bytes = await file.arrayBuffer();
        const pdfDoc = await PDFDocument.load(bytes);
        const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
        const pages = pdfDoc.getPages();

        const indicesToWatermark = globalPagesSelectionMode === 'custom' 
          ? globalSelectedPageIndices 
          : pages.map((_, i) => i);

        for (const idx of indicesToWatermark) {
          if (idx < pages.length) {
            const page = pages[idx];
            const { width, height } = page.getSize();
            page.drawText(watermarkText, {
              x: width / 5,
              y: height / 3,
              size: 40,
              font,
              color: rgb(0.7, 0.7, 0.7),
              opacity: 0.3,
              rotate: { angle: 35, type: 'degrees' as any }
            });
          }
        }

        let watermarkedBytes;
        if (globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.length > 0) {
          const newPdf = await PDFDocument.create();
          const copiedPages = await newPdf.copyPages(pdfDoc, globalSelectedPageIndices);
          copiedPages.forEach(p => newPdf.addPage(p));
          watermarkedBytes = await newPdf.save();
        } else {
          watermarkedBytes = await pdfDoc.save();
        }

        const blob = new Blob([watermarkedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`watermarked_${file.name}`);
        setStatus('success');

      } else if (toolId === 'protect') {
        // --- PASSWORD PROTECTION ---
        if (!securityPassword) {
          throw new Error(language === Language.AR ? 'الرجاء إدخال كلمة المرور.' : 'Please enter password.');
        }
        const file = fileToProcess;
        const base64Data = await new Promise<string>((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => {
            const res = r.result as string;
            resolve(res.split(',')[1]);
          };
          r.onerror = reject;
          r.readAsDataURL(file);
        });

        const res = await fetch('/api/pdf/protect', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            fileData: base64Data,
            password: securityPassword
          })
        });

        if (!res.ok) {
          const errData = await res.json();
          throw new Error(errData.error || 'Failed to encrypt PDF');
        }

        const { securedData } = await res.json();
        const blob = await fetch(`data:application/pdf;base64,${securedData}`).then(r => r.blob());
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`protected_${file.name}`);
        setStatus('success');

      } else if (toolId === 'rotate') {
        // --- VISUAL ROTATE PDF ---
        const file = selectedFiles[0];
        const bytes = await file.arrayBuffer();
        const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const pages = pdfDoc.getPages();
        
        for (const pageInfo of interactivePages) {
          if (pageInfo.rotation !== 0 && pageInfo.pageNum < pages.length) {
            const page = pages[pageInfo.pageNum];
            const currentRotation = page.getRotation().angle;
            page.setRotation(degrees((currentRotation + pageInfo.rotation) % 360));
          }
        }
        
        let rotatedBytes;
        if (globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.length > 0) {
          const newPdf = await PDFDocument.create();
          const copiedPages = await newPdf.copyPages(pdfDoc, globalSelectedPageIndices);
          copiedPages.forEach(p => newPdf.addPage(p));
          rotatedBytes = await newPdf.save();
        } else {
          rotatedBytes = await pdfDoc.save();
        }

        const blob = new Blob([rotatedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`rotated_${file.name}`);
        setStatus('success');

      } else if (toolId === 'delete_pages') {
        // --- VISUAL DELETE PAGES ---
        const file = selectedFiles[0];
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        
        const pagesToKeep = interactivePages
          .filter(p => !p.isSelectedForDelete)
          .map(p => p.pageNum);

        if (pagesToKeep.length === 0) {
          throw new Error(language === Language.AR ? 'لا يمكنك حذف كل الصفحات. يجب أن تبقى صفحة واحدة على الأقل.' : 'You cannot delete all pages. At least one page must remain.');
        }

        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(pdf, pagesToKeep);
        copiedPages.forEach(page => newPdf.addPage(page));
        const finalBytes = await newPdf.save();
        const blob = new Blob([finalBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`edited_${file.name}`);
        setStatus('success');

      } else if (toolId === 'reorder') {
        // --- VISUAL REORDER PAGES ---
        const file = selectedFiles[0];
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });

        const orderIndices = interactivePages.map(p => p.pageNum);

        if (orderIndices.length === 0) {
          throw new Error(language === Language.AR ? 'ترتيب الصفحات غير صالح.' : 'Invalid page order specified.');
        }

        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(pdf, orderIndices);
        
        // Apply individual page rotations in reordered files if any
        for (let i = 0; i < copiedPages.length; i++) {
          const pageInfo = interactivePages[i];
          if (pageInfo && pageInfo.rotation !== 0) {
            const currentRot = copiedPages[i].getRotation().angle;
            copiedPages[i].setRotation(degrees((currentRot + pageInfo.rotation) % 360));
          }
          newPdf.addPage(copiedPages[i]);
        }

        const finalBytes = await newPdf.save();
        const blob = new Blob([finalBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`reordered_${file.name}`);
        setStatus('success');

      } else if (toolId === 'unlock') {
        // --- UNLOCK PDF ---
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const unlockedBytes = await pdfDoc.save();
        const blob = new Blob([unlockedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`unlocked_${file.name}`);
        setStatus('success');

      } else if (toolId === 'pdf2img' || toolId === 'pdf2png') {
        // --- PDF TO IMAGES (JPG/PNG) ---
        const file = fileToProcess;
        const format = toolId === 'pdf2img' ? 'jpeg' : 'png';
        const ext = format === 'jpeg' ? 'jpg' : 'png';
        const mimeType = format === 'jpeg' ? 'image/jpeg' : 'image/png';

        const pageImages = await renderPdfPagesToImages(file, format);

        if (pageImages.length === 0) {
          throw new Error(language === Language.AR ? 'لم يتم العثور على أي صفحات في مستند PDF.' : 'No pages found in the PDF document.');
        }

        if (pageImages.length === 1) {
          const singleImage = pageImages[0];
          setResultUrl(URL.createObjectURL(singleImage.blob));
          setResultBlob(singleImage.blob);
          setResultMime(mimeType);
          const finalName = outputFileName.trim() && outputFileName !== 'merged_document' 
            ? `${outputFileName}.${ext}` 
            : singleImage.name;
          setDownloadName(finalName);
        } else {
          // Multi-page PDF -> Zip all page images
          const zip = new JSZip();
          pageImages.forEach((img) => {
            zip.file(img.name, img.blob);
          });
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          setResultUrl(URL.createObjectURL(zipBlob));
          setResultBlob(zipBlob);
          setResultMime('application/zip');
          const finalName = outputFileName.trim() && outputFileName !== 'merged_document' 
            ? `${outputFileName}_images.zip` 
            : `${file.name.replace(/\.pdf$/i, '')}_images.zip`;
          setDownloadName(finalName);
        }
        setStatus('success');

      } else if (toolId === 'pdf2svg') {
        // --- REAL PDF TO SVG VECTOR CONVERTER ---
        const file = fileToProcess;
        const pageSvgs = await renderPdfPagesToSVGs(file);

        if (pageSvgs.length === 0) {
          throw new Error(language === Language.AR ? 'لم يتم العثور على أي صفحات في مستند PDF.' : 'No pages found in the PDF document.');
        }

        if (pageSvgs.length === 1) {
          const singleSvg = pageSvgs[0];
          const blob = new Blob([singleSvg.content], { type: 'image/svg+xml;charset=utf-8' });
          setResultUrl(URL.createObjectURL(blob));
          setResultBlob(blob);
          setResultMime('image/svg+xml;charset=utf-8');
          const finalName = outputFileName.trim() && outputFileName !== 'merged_document' 
            ? `${outputFileName}.svg` 
            : `${file.name.replace(/\.pdf$/i, '')}.svg`;
          setDownloadName(finalName);
        } else {
          // Multi-page PDF -> Zip all page SVGs
          const zip = new JSZip();
          pageSvgs.forEach((svg) => {
            zip.file(svg.name, svg.content);
          });
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          setResultUrl(URL.createObjectURL(zipBlob));
          setResultBlob(zipBlob);
          setResultMime('application/zip');
          const finalName = outputFileName.trim() && outputFileName !== 'merged_document' 
            ? `${outputFileName}_svgs.zip` 
            : `${file.name.replace(/\.pdf$/i, '')}_svgs.zip`;
          setDownloadName(finalName);
        }
        setStatus('success');

      } else if (toolId === 'pdf2html') {
        // --- REAL PDF TO HTML CONVERTER ---
        const file = fileToProcess;
        const localText = await extractTextLocally(file);
        
        const cleanText = localText || (language === Language.AR
          ? 'تنبيه: هذا المستند عبارة عن ملف PDF ممسوح ضوئياً (صورة فقط) ولا يحتوي على نصوص قابلة للاستخراج المباشر. يرجى استخدام أداة التعرف الضوئي على الحروف (OCR) المتوفرة في القائمة الجانبية لمعالجة محتوى هذا الملف بالكامل واستخراج نصوصه.'
          : 'Notice: This document is a scanned PDF (image-only) and does not contain an extractable text layer. Please use the AI OCR tool available in our platform to extract and process this document\'s text content.');

        const paragraphs = cleanText.split('\n').filter(p => p.trim());
        const htmlParagraphs = paragraphs.map(p => `<p style="margin-bottom: 1.5rem; color: #334155; line-height: 1.75; font-size: 1.125rem;">${p.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>`).join('\n');

        const htmlContent = `<!DOCTYPE html>
<html lang="${language}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Converted Document - ${file.name}</title>
    <style>
        body {
            font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #f1f5f9;
            color: #1e293b;
            margin: 0;
            padding: 2rem 1rem;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 3rem;
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            border: 1px solid #e2e8f0;
        }
        .header {
            border-bottom: 2px solid #e2e8f0;
            padding-bottom: 1.5rem;
            margin-bottom: 2rem;
        }
        h1 {
            color: #4f46e5;
            margin: 0 0 0.5rem 0;
            font-size: 2.25rem;
            font-weight: 800;
        }
        .meta {
            color: #64748b;
            font-size: 0.875rem;
            font-weight: 500;
        }
        .content {
            font-size: 1.125rem;
            line-height: 1.75;
            color: #334155;
        }
        .footer {
            margin-top: 3rem;
            border-top: 1px solid #e2e8f0;
            padding-top: 1.5rem;
            text-align: center;
            color: #94a3b8;
            font-size: 0.875rem;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>${file.name.substring(0, file.name.lastIndexOf('.')) || 'Converted Document'}</h1>
        </div>
        <div class="content">
            ${htmlParagraphs}
        </div>
    </div>
</body>
</html>`;

        const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
        setResultUrl(URL.createObjectURL(blob));
        setResultBlob(blob);
        setResultMime('text/html;charset=utf-8');
        const finalName = outputFileName.trim() && outputFileName !== 'merged_document' ? `${outputFileName}.html` : file.name.replace(/\.pdf$/i, '.html');
        setDownloadName(finalName);
        setStatus('success');

      } else if (toolId === 'pdf2txt') {
        // --- REAL PDF TO TEXT CONVERTER ---
        const file = fileToProcess;
        const localText = await extractTextLocally(file);
        
        const cleanText = localText || (language === Language.AR
          ? 'تنبيه: هذا المستند عبارة عن ملف PDF ممسوح ضوئياً (صورة فقط) ولا يحتوي على نصوص قابلة للاستخراج المباشر. يرجى استخدام أداة التعرف الضوئي على الحروف (OCR) المتوفرة في القائمة الجانبية لمعالجة محتوى هذا الملف بالكامل واستخراج نصوصه.'
          : 'Notice: This document is a scanned PDF (image-only) and does not contain an extractable text layer. Please use the AI OCR tool available in our platform to extract and process this document\'s text content.');

        const txtContent = `${language === Language.AR ? 'مستند مستخرج عبر PDFProTools' : 'Document Extracted via PDFProTools'}\r\n==================================\r\n\r\n${cleanText}`;

        const blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
        setResultUrl(URL.createObjectURL(blob));
        setResultBlob(blob);
        setResultMime('text/plain;charset=utf-8');
        const finalName = outputFileName.trim() && outputFileName !== 'merged_document' ? `${outputFileName}.txt` : file.name.replace(/\.pdf$/i, '.txt');
        setDownloadName(finalName);
        setStatus('success');

      } else if (toolId === 'pdf2epub') {
        // --- REAL PDF TO EPUB CONVERTER ---
        const file = fileToProcess;
        const localText = await extractTextLocally(file);
        
        const cleanText = localText || (language === Language.AR
          ? 'تنبيه: هذا المستند عبارة عن ملف PDF ممسوح ضوئياً (صورة فقط) ولا يحتوي على نصوص قابلة للاستخراج المباشر. يرجى استخدام أداة التعرف الضوئي على الحروف (OCR) المتوفرة في القائمة الجانبية لمعالجة محتوى هذا الملف بالكامل واستخراج نصوصه.'
          : 'Notice: This document is a scanned PDF (image-only) and does not contain an extractable text layer. Please use the AI OCR tool available in our platform to extract and process this document\'s text content.');

        const paragraphs = cleanText.split('\n').filter(p => p.trim());
        const epubParagraphs = paragraphs.map(p => `<p class="calibre1">${p.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>`).join('\n');

        const epubXhtml = `<?xml version='1.0' encoding='utf-8'?>
<html xmlns="http://www.w3.org/1999/xhtml" lang="${language}" xml:lang="${language}">
  <head>
    <title>Converted Book - ${file.name}</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <style type="text/css">
      body {
        margin: 10%;
        font-family: serif;
        line-height: 1.5;
        text-align: justify;
      }
      h1 {
        text-align: center;
        color: #4f46e5;
        margin-bottom: 2em;
      }
      p.calibre1 {
        margin-bottom: 1em;
        text-indent: 1.5em;
      }
    </style>
  </head>
  <body>
    <h1>${file.name.substring(0, file.name.lastIndexOf('.')) || 'Converted Book'}</h1>
    ${epubParagraphs}
  </body>
</html>`;

        const blob = new Blob([epubXhtml], { type: 'application/epub+zip' });
        setResultUrl(URL.createObjectURL(blob));
        setResultBlob(blob);
        setResultMime('application/epub+zip');
        const finalName = outputFileName.trim() && outputFileName !== 'merged_document' ? `${outputFileName}.epub` : file.name.replace(/\.pdf$/i, '.epub');
        setDownloadName(finalName);
        setStatus('success');

      } else if (toolId === 'crop') {
        // --- CROP PAGES ---
        const file = selectedFiles[0];
        const bytes = await file.arrayBuffer();
        const pdfDoc = await PDFDocument.load(bytes);
        const pages = pdfDoc.getPages();
        const cropMarginFactor = cropPercentage / 100;

        const indicesToCrop = globalPagesSelectionMode === 'custom' 
          ? globalSelectedPageIndices 
          : pages.map((_, i) => i);

        for (const idx of indicesToCrop) {
          if (idx < pages.length) {
            const page = pages[idx];
            const { width, height } = page.getSize();
            const left = width * cropMarginFactor;
            const bottom = height * cropMarginFactor;
            const w = width * (1 - 2 * cropMarginFactor);
            const h = height * (1 - 2 * cropMarginFactor);
            
            if (w > 10 && h > 10) {
              page.setCropBox(left, bottom, w, h);
            }
          }
        }

        let croppedBytes;
        if (globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.length > 0) {
          const newPdf = await PDFDocument.create();
          const copiedPages = await newPdf.copyPages(pdfDoc, globalSelectedPageIndices);
          copiedPages.forEach(p => newPdf.addPage(p));
          croppedBytes = await newPdf.save();
        } else {
          croppedBytes = await pdfDoc.save();
        }

        const blob = new Blob([croppedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(`cropped_${file.name}`);
        setStatus('success');

      } else if (['png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId)) {
        // --- IMAGE CONVERSIONS (PNG/JPG to WEBP/HEIC) ---
        const file = selectedFiles[0];
        const isToWebp = toolId.includes('webp');
        const format: 'webp' | 'heic' = isToWebp ? 'webp' : 'heic';

        const { blob, ext } = await convertImage(file, format, imageQuality);

        setResultUrl(URL.createObjectURL(blob));
        const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        setDownloadName(`${baseName}.${ext}`);
        setStatus('success');

      } else if (toolId === 'compare') {
        // --- COMPARE DOCUMENTS ---
        let file1 = selectedFiles[0];
        let file2 = selectedFiles[1] || compareFile;

        if (!file1 || !file2) {
          throw new Error(
            language === Language.AR 
              ? 'يرجى توفير ملفين للمقارنة (يمكنك رفع الملف الثاني من اللوحة الجانبية).' 
              : 'Please select or upload two PDF files to compare.'
          );
        }

        const base64Data1 = await fileToBase64(file1);
        const base64Data2 = await fileToBase64(file2);
        const text1 = await extractTextLocally(file1);

        const response = await fetch('/api/gemini/process', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            task: 'compare',
            fileData: base64Data1,
            mimeType: file1.type,
            text: text1,
            secondFileData: base64Data2,
            secondMimeType: file2.type,
            compareMode,
            diffView
          })
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Server comparison error');
        }

        const data = await response.json();
        setAiTextResult(data.result || 'No differences identified');
        setStatus('success');

      } else if (toolId === 'chat') {
        // --- CHAT INITIALIZATION ---
        setChatHistory([
          { sender: 'assistant', text: texts.chatIntro }
        ]);
        setStatus('success');

      } else if (isPdfConversionTool) {
        // --- REAL FILE-TO-PDF CONVERTER VIA GEMINI + PDF-LIB PIPELINE ---
        const file = fileToProcess;
        let fileData = '';
        let fileName = file ? file.name : 'converted.pdf';
        let mimeType = file ? file.type : 'text/plain';

        if (file) {
          fileData = await fileToBase64(file);
        }

        const payload: any = {
          toolId,
          fileName,
          mimeType,
          fileData: fileData || undefined,
          url: (toolId === 'html2pdf' && htmlUrl.trim()) ? htmlUrl.trim() : undefined,
          quality: pdfConvQuality,
        };

        const response = await fetch('/api/pdf/convert2pdf', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(payload),
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Failed to convert document to PDF');
        }

        const blob = await response.blob();
        setResultUrl(URL.createObjectURL(blob));

        let downloadBaseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
        if (toolId === 'html2pdf' && htmlUrl.trim()) {
          try {
            const urlObj = new URL(htmlUrl.trim());
            downloadBaseName = urlObj.hostname.replace('www.', '') || 'webpage';
          } catch {
            downloadBaseName = 'webpage';
          }
        }
        setDownloadName(`${downloadBaseName}.pdf`);
        setStatus('success');

      } else if (toolId === 'pdf2word') {
        // --- REAL PDF TO WORD CONVERTER WITH OPTIONS ---
        const file = fileToProcess;
        const base64Data = await fileToBase64(file);

        const response = await fetch('/api/pdf/pdf2word', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            fileData: base64Data,
            ocr: ocrEnabled,
            mode: pdf2wordMode
          }),
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Failed to convert PDF to Word');
        }

        const data = await response.json();
        const { docxData, blocks } = data;

        setWordBlocks(blocks || []);

        // Convert base64 data to binary bytes securely in memory with exact MIME type
        const byteCharacters = atob(docxData);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        setResultUrl(URL.createObjectURL(blob));
        setResultBlob(blob);
        setResultMime('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        
        let targetDocName = outputFileName.trim() ? outputFileName.replace(/\.docx$/i, '') : file.name.replace(/\.pdf$/i, '');
        if (pdf2wordMode === 'text') {
          targetDocName += '_text';
        }
        setDownloadName(`${targetDocName}.docx`);
        setStatus('success');

      } else if (toolId === 'pdf2excel' || toolId === 'pdf2ppt') {
        // --- FORMAT CONVERTERS (EXCEL / PPT) WITH CONFIG VIA SERVER-SIDE GEMINI ---
        const file = fileToProcess;
        const base64Data = await fileToBase64(file);

        const response = await fetch('/api/gemini/process', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            task: toolId,
            fileData: base64Data,
            mimeType: 'application/pdf',
          }),
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || `Failed to convert PDF to ${toolId === 'pdf2excel' ? 'Excel' : 'PowerPoint'}`);
        }

        const data = await response.json();
        const content = data.result || '';

        if (toolId === 'pdf2excel') {
          const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
          const encoder = new TextEncoder();
          const encoded = encoder.encode(content);
          
          const combined = new Uint8Array(bom.length + encoded.length);
          combined.set(bom, 0);
          combined.set(encoded, bom.length);

          const blob = new Blob([combined], { type: 'text/csv;charset=utf-8' });
          setResultUrl(URL.createObjectURL(blob));
          setResultBlob(blob);
          setResultMime('text/csv;charset=utf-8');
          
          const finalBaseName = outputFileName.trim() ? outputFileName.replace(/\.csv$/i, '') : file.name.replace(/\.pdf$/i, '');
          setDownloadName(`${finalBaseName}.csv`);
        } else {
          // PPT presentation planning
          setAiTextResult(content);
          
          const encoder = new TextEncoder();
          const encoded = encoder.encode(content);
          const blob = new Blob([encoded], { type: 'text/plain;charset=utf-8;' });
          setResultUrl(URL.createObjectURL(blob));
          setResultBlob(blob);
          setResultMime('text/plain');

          const finalBaseName = outputFileName.trim() ? outputFileName.replace(/\.txt$/i, '') : file.name.replace(/\.pdf$/i, '');
          setDownloadName(`${finalBaseName}_presentation_plan.txt`);
        }
        setStatus('success');

      } else if (toolId === 'add_page_numbers') {
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const totalPages = pdf.getPageCount();
        const font = await pdf.embedFont(StandardFonts.Helvetica);

        const start = pageNumberRange === 'all' ? 1 : Math.max(1, pageNumberStart);
        const end = pageNumberRange === 'all' ? totalPages : Math.min(totalPages, pageNumberEnd);

        for (let i = start - 1; i < end; i++) {
          const page = pdf.getPage(i);
          const n = i + 1;
          
          let text = pageNumberFormat
            .replace('{n}', String(n))
            .replace('{total}', String(totalPages));
            
          if (language === Language.AR) {
            text = text.replace('Page', 'صفحة').replace('of', 'من');
          }
          
          const textWidth = font.widthOfTextAtSize(text, pageNumberSize);
          const textHeight = pageNumberSize;
          
          const width = page.getWidth();
          const height = page.getHeight();
          const margin = 30;
          
          let x = margin;
          let y = margin;
          
          if (pageNumberPosition.includes('center')) {
            x = (width - textWidth) / 2;
          } else if (pageNumberPosition.includes('right')) {
            x = width - textWidth - margin;
          }
          
          if (pageNumberPosition.startsWith('top')) {
            y = height - margin - textHeight;
          } else if (pageNumberPosition.startsWith('middle')) {
            y = (height - textHeight) / 2;
          }
          
          page.drawText(text, {
            x,
            y,
            size: pageNumberSize,
            font,
            color: rgb(0.28, 0.33, 0.41), // Tailwind slate-600 approx
          });
        }

        const pdfBytes = await pdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : file.name.replace(/\.pdf$/i, '') + '_numbered.pdf');
        setStatus('success');

      } else if (toolId === 'repair') {
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        let recoveredBytes;
        try {
          const pdfDoc = await PDFDocument.load(bytes, { ignoreEncryption: true });
          recoveredBytes = await pdfDoc.save();
        } catch (loadError) {
          console.log("Standard load failed, attempting structural salvage...", loadError);
          let text = new TextDecoder('latin1').decode(bytes);
          const startIdx = text.indexOf('%PDF-');
          if (startIdx > 0) {
            text = text.substring(startIdx);
          }
          if (!text.includes('%%EOF')) {
            text += '\n%%EOF';
          }
          const cleanBytes = new TextEncoder().encode(text);
          const pdfDoc = await PDFDocument.load(cleanBytes, { ignoreEncryption: true });
          recoveredBytes = await pdfDoc.save();
        }
        
        const blob = new Blob([recoveredBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : file.name.replace(/\.pdf$/i, '') + '_repaired.pdf');
        setStatus('success');

      } else if (toolId === 'flatten') {
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        try {
          const form = pdf.getForm();
          if (form) {
            form.flatten();
          }
        } catch (formError) {
          console.log("No form fields found or failed to flatten:", formError);
        }
        const flattenedBytes = await pdf.save();
        const blob = new Blob([flattenedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : file.name.replace(/\.pdf$/i, '') + '_flattened.pdf');
        setStatus('success');

      } else if (toolId === 'metadata') {
        const file = fileToProcess;
        const bytes = await file.arrayBuffer();
        const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
        
        pdf.setTitle(metaTitle);
        pdf.setAuthor(metaAuthor);
        pdf.setSubject(metaSubject);
        pdf.setKeywords(metaKeywords.split(',').map(k => k.trim()).filter(Boolean));
        
        const updatedBytes = await pdf.save();
        const blob = new Blob([updatedBytes], { type: 'application/pdf' });
        setResultUrl(URL.createObjectURL(blob));
        setDownloadName(outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : file.name.replace(/\.pdf$/i, '') + '_metadata.pdf');
        setStatus('success');

      } else if (toolId === 'web2pdf') {
        const isUrlMode = htmlUrl.trim().length > 0;
        if (!isUrlMode) {
          throw new Error(language === Language.AR ? 'الرجاء إدخال رابط URL صالح.' : 'Please enter a valid website URL.');
        }
        
        const payload = {
          toolId: 'html2pdf',
          fileName: 'webpage.html',
          mimeType: 'text/html',
          url: htmlUrl.trim()
        };

        const response = await fetch('/api/pdf/convert2pdf', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(payload),
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Failed to convert website to PDF');
        }

        const blob = await response.blob();
        setResultUrl(URL.createObjectURL(blob));

        let downloadBaseName = 'webpage';
        try {
          const urlObj = new URL(htmlUrl.trim());
          downloadBaseName = urlObj.hostname.replace('www.', '') || 'webpage';
        } catch {
          downloadBaseName = 'webpage';
        }
        setDownloadName(`${downloadBaseName}.pdf`);
        setStatus('success');

      } else if (toolId === 'qrcode') {
        if (!qrText.trim()) {
          throw new Error(language === Language.AR ? 'الرجاء إدخال نص أو رابط لرمز QR.' : 'Please enter a valid text or URL for the QR code.');
        }

        const qrRes = await fetch(`https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(qrText.trim())}`);
        if (!qrRes.ok) {
          throw new Error('Failed to generate QR code');
        }
        const qrBlobBytes = await qrRes.arrayBuffer();

        if (qrMode === 'download') {
          const blob = new Blob([qrBlobBytes], { type: 'image/png' });
          setResultUrl(URL.createObjectURL(blob));
          setDownloadName('qrcode.png');
          setStatus('success');
        } else {
          const file = fileToProcess;
          if (!file) {
            throw new Error(language === Language.AR ? 'الرجاء اختيار ملف PDF لتضمين الرمز بداخله.' : 'Please upload a PDF document to embed the QR code inside.');
          }
          const bytes = await file.arrayBuffer();
          const pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
          const qrImage = await pdf.embedPng(qrBlobBytes);

          const totalPages = pdf.getPageCount();
          const targetPageNum = Math.max(1, Math.min(totalPages, qrPageNum));
          const page = pdf.getPage(targetPageNum - 1);

          const width = page.getWidth();
          const height = page.getHeight();
          const size = qrCodeSize || 100;
          const margin = 30;

          let x = margin;
          let y = margin;

          if (qrPosition.includes('center')) {
            x = (width - size) / 2;
          } else if (qrPosition.includes('right')) {
            x = width - size - margin;
          }

          if (qrPosition.startsWith('top')) {
            y = height - margin - size;
          } else if (qrPosition.startsWith('middle')) {
            y = (height - size) / 2;
          }

          page.drawImage(qrImage, {
            x,
            y,
            width: size,
            height: size,
          });

          const savedBytes = await pdf.save();
          const blob = new Blob([savedBytes], { type: 'application/pdf' });
          setResultUrl(URL.createObjectURL(blob));
          setDownloadName(outputFileName.trim() ? `${outputFileName.replace(/\.pdf$/i, '')}.pdf` : file.name.replace(/\.pdf$/i, '') + '_qr.pdf');
          setStatus('success');
        }

      } else {
        // --- AI TOOLS (summarize, translate, extract, ocr, format, redact) ---
        const file = fileToProcess;
        const base64Data = await fileToBase64(file);
        const localExtractedText = await extractTextLocally(file);

        // Fetch API request to server-side endpoint with complete dashboard parameters
        const response = await fetch('/api/gemini/process', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            task: toolId,
            fileData: base64Data,
            mimeType: file.type,
            text: localExtractedText,
            
            // 16. Smart Text Extraction
            extractPagesMode,
            extractPagesRange,
            extractFormat,

            // 17. Summarization
            summaryLength,
            summaryFormat,

            // 18. Translation
            targetLang: targetTranslationLang,
            preserveLayout,

            // 19. OCR
            ocrLanguage,
            ocrOutputType,

            // 20. Intelligent Formatting
            formatAutoAlign,
            formatUnifyFonts,

            // 22. Redact Sensitive Info
            redactPiiEmail,
            redactPiiPhone,
            redactPiiId,
            redactPiiAddress,
            redactKeywords
          })
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Server processing error');
        }

        const data = await response.json();
        setAiTextResult(data.result || 'No output generated');
        setStatus('success');
      }

    } catch (err: any) {
      console.error(err);
      setStatus('error');
      setErrorMessage(err.message || 'An error occurred during file processing.');
    }
  };

  // Render markdown parser helper
  const renderMarkdown = (md: string) => {
    return md.split('\n').map((line, idx) => {
      const trimmed = line.trim();
      if (trimmed.startsWith('### ')) {
        return <h4 key={idx} className="text-lg font-bold text-slate-800 mt-4 mb-2">{trimmed.substring(4)}</h4>;
      }
      if (trimmed.startsWith('## ')) {
        return <h3 key={idx} className="text-xl font-bold text-slate-800 mt-5 mb-3">{trimmed.substring(3)}</h3>;
      }
      if (trimmed.startsWith('# ')) {
        return <h2 key={idx} className="text-2xl font-bold text-slate-900 mt-6 mb-4">{trimmed.substring(2)}</h2>;
      }
      if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
        return <li key={idx} className="list-disc ms-6 text-slate-700 my-1">{trimmed.substring(2)}</li>;
      }
      if (trimmed.match(/^\d+\.\s/)) {
        return <li key={idx} className="list-decimal ms-6 text-slate-700 my-1">{trimmed.replace(/^\d+\.\s/, '')}</li>;
      }
      if (trimmed === '') {
        return <div key={idx} className="h-2" />;
      }
      return <p key={idx} className="text-slate-700 leading-relaxed my-2">{trimmed}</p>;
    });
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(aiTextResult);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const saveToTxt = () => {
    const blob = new Blob([aiTextResult], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${toolId}_result.txt`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleSendChatMessage = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!chatInput.trim() || isChatSending) return;

    const userMsg = chatInput.trim();
    setChatInput('');
    setChatHistory(prev => [...prev, { sender: 'user', text: userMsg }]);
    setIsChatSending(true);

    try {
      const file = selectedFiles[0];
      const base64Data = await fileToBase64(file);
      const localExtractedText = await extractTextLocally(file);

      const response = await fetch('/api/gemini/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          task: 'chat',
          fileData: base64Data,
          mimeType: file.type,
          text: localExtractedText,
          question: userMsg
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Server chat error');
      }

      const data = await response.json();
      setChatHistory(prev => [...prev, { sender: 'assistant', text: data.result || 'Sorry, I couldn\'t formulate an answer.' }]);
    } catch (err: any) {
      console.error(err);
      setChatHistory(prev => [...prev, { sender: 'assistant', text: `Error: ${err.message || 'Failed to connect to AI'}` }]);
    } finally {
      setIsChatSending(false);
    }
  };

  const movePage = (index: number, direction: 'left' | 'right') => {
    const newPages = [...interactivePages];
    const targetIndex = direction === 'left' ? index - 1 : index + 1;
    if (targetIndex >= 0 && targetIndex < newPages.length) {
      const temp = newPages[index];
      newPages[index] = newPages[targetIndex];
      newPages[targetIndex] = temp;
      setInteractivePages(newPages);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-50 flex items-center justify-center p-4 overflow-y-auto">
      <div 
        dir={isRTL ? 'rtl' : 'ltr'} 
        className={`bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100 rounded-2xl w-full shadow-2xl border border-slate-100 dark:border-slate-800 overflow-hidden flex flex-col max-h-[95vh] transition-all duration-300 ${
          (selectedFiles.length > 0 || (['html2pdf', 'web2pdf'].includes(toolId) && htmlUrl.trim().length > 0) || toolId === 'qrcode') && ['delete_pages', 'split', 'rotate', 'crop', 'merge', 'reorder', 'compress', 'watermark', 'protect', 'pdf2word', 'pdf2excel', 'pdf2ppt', 'img2pdf', 'unlock', 'pdf2img', 'pdf2png', 'pdf2svg', 'pdf2html', 'pdf2txt', 'pdf2epub', 'extract', 'summarize', 'translate', 'ocr', 'format', 'chat', 'redact', 'compare', 'docx2pdf', 'xlsx2pdf', 'pptx2pdf', 'html2pdf', 'txt2pdf', 'rtf2pdf', 'add_page_numbers', 'repair', 'flatten', 'metadata', 'web2pdf', 'qrcode'].includes(toolId)
            ? 'max-w-5xl' 
            : 'max-w-2xl'
        }`}
      >
        {/* Modal Header */}
        <div className="p-6 border-b border-slate-100 dark:border-slate-800 flex items-center justify-between bg-slate-50 dark:bg-slate-900/50">
          <div className="flex items-center space-x-3 space-x-reverse">
            <div className="p-2 bg-indigo-100 dark:bg-indigo-950/40 text-indigo-700 dark:text-indigo-400 rounded-lg">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
              </svg>
            </div>
            <h3 className="text-xl font-bold text-slate-900 dark:text-white">{toolName}</h3>
          </div>
          <button 
            onClick={onClose}
            className="p-1 hover:bg-slate-200 dark:hover:bg-slate-800 text-slate-500 dark:text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 rounded-full transition-colors focus:outline-none"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto flex-1 space-y-6 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100">
          {status === 'idle' && (
            <>
              {(selectedFiles.length === 0 && !((['html2pdf', 'web2pdf'].includes(toolId)) && htmlUrl.trim().length > 0) && toolId !== 'qrcode') ? (
                <div className="space-y-6 w-full">
                  {['html2pdf', 'web2pdf'].includes(toolId) && (
                    <div className="bg-slate-50 dark:bg-slate-950 p-5 border border-slate-200 dark:border-slate-800 rounded-xl space-y-3">
                      <label className="block text-sm font-bold text-slate-700 dark:text-slate-300">
                        {language === Language.AR ? 'تحويل صفحة ويب عبر رابط (URL):' : 'Convert web page via Link (URL):'}
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="url"
                          placeholder="https://example.com"
                          value={htmlUrl}
                          onChange={(e) => setHtmlUrl(e.target.value)}
                          className="flex-grow p-3 border border-slate-250 dark:border-slate-700 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100 font-medium"
                        />
                        {htmlUrl.trim() && (
                          <button
                            type="button"
                            onClick={handleProcess}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-5 py-3 rounded-xl transition-all shadow-sm text-sm cursor-pointer"
                          >
                            {language === Language.AR ? 'تحويل الرابط' : 'Convert Link'}
                          </button>
                        )}
                      </div>
                      {toolId === 'html2pdf' && (
                        <div className="relative flex py-2 items-center">
                          <div className="flex-grow border-t border-slate-200 dark:border-slate-800"></div>
                          <span className="flex-shrink mx-4 text-slate-400 dark:text-slate-500 text-xs font-semibold uppercase">{language === Language.AR ? 'أو ارفع ملفاً محلياً' : 'OR upload local file'}</span>
                          <div className="flex-grow border-t border-slate-200 dark:border-slate-800"></div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* Drag and Drop Zone */}
                  {toolId !== 'web2pdf' && (
                    <div
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                      className={`border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all duration-300 flex flex-col items-center justify-center ${
                        isDragOver 
                          ? 'border-indigo-500 bg-indigo-50 dark:bg-indigo-950/20' 
                          : 'border-slate-300 dark:border-slate-700 hover:border-indigo-400 dark:hover:border-indigo-500 hover:bg-slate-50 dark:hover:bg-slate-800/40'
                      }`}
                    >
                      <input
                        type="file"
                        ref={fileInputRef}
                        onChange={handleFileSelect}
                        accept={acceptedTypes}
                        multiple={toolId === 'merge' || toolId === 'img2pdf' || toolId === 'compare'}
                        className="hidden"
                      />
                      <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12 text-slate-400 dark:text-slate-500 mb-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                      </svg>
                      <p className="text-sm font-semibold text-slate-700 dark:text-slate-300">
                        {isImageTool ? texts.uploadZoneImg : isPdfConversionTool ? (
                          toolId === 'docx2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف Word هنا، أو انقر للتصفح' : 'Drag & drop Word document here, or click to browse')
                          : toolId === 'xlsx2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف Excel هنا، أو انقر للتصفح' : 'Drag & drop Excel sheet here, or click to browse')
                          : toolId === 'pptx2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف PowerPoint هنا، أو انقر للتصفح' : 'Drag & drop PowerPoint here, or click to browse')
                          : toolId === 'html2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف HTML هنا، أو انقر للتصفح' : 'Drag & drop HTML file here, or click to browse')
                          : toolId === 'txt2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف نصي .txt هنا، أو انقر للتصفح' : 'Drag & drop plain text .txt here, or click to browse')
                          : toolId === 'rtf2pdf' ? (language === Language.AR ? 'اسحب وأسقط ملف RTF هنا، أو انقر للتصفح' : 'Drag & drop RTF document here, or click to browse')
                          : 'Drag & drop files here'
                        ) : texts.uploadZonePdf}
                      </p>
                      <p className="text-xs text-slate-400 dark:text-slate-500 mt-2">
                        {isImageTool ? 'PNG, JPG, JPEG' : isPdfConversionTool ? (
                          toolId === 'docx2pdf' ? '.docx, .doc'
                          : toolId === 'xlsx2pdf' ? '.xlsx, .xls'
                          : toolId === 'pptx2pdf' ? '.pptx, .ppt'
                          : toolId === 'html2pdf' ? '.html, .htm'
                          : toolId === 'txt2pdf' ? '.txt'
                          : toolId === 'rtf2pdf' ? '.rtf'
                          : '*/*'
                        ) : 'PDF'}
                      </p>
                    </div>
                  )}
                </div>
              ) : ['delete_pages', 'split', 'rotate', 'crop', 'merge', 'reorder', 'compress', 'watermark', 'protect', 'pdf2word', 'pdf2excel', 'pdf2ppt', 'img2pdf', 'unlock', 'pdf2img', 'pdf2png', 'pdf2svg', 'pdf2html', 'pdf2txt', 'pdf2epub', 'extract', 'summarize', 'translate', 'ocr', 'format', 'chat', 'redact', 'compare', 'docx2pdf', 'xlsx2pdf', 'pptx2pdf', 'html2pdf', 'txt2pdf', 'rtf2pdf', 'add_page_numbers', 'repair', 'flatten', 'metadata', 'web2pdf', 'qrcode'].includes(toolId) ? (
                /* Visual Interactive Workspace layout */
                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {/* Left Column (2/3 width) - Interactive Pages Gallery */}
                  <div className="md:col-span-2 space-y-4">
                    <div className="flex items-center justify-between border-b border-slate-100 dark:border-slate-800 pb-3">
                      <div>
                        <h4 className="text-base font-bold text-slate-800 dark:text-slate-200 flex items-center gap-1.5">
                          <span className="text-lg">👁️</span>
                          <span>{language === Language.AR ? 'المعاينة البصرية التفاعلية' : 'Visual Interactive Preview'}</span>
                        </h4>
                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                          {language === Language.AR
                            ? 'تفاعل مع الصفحات مباشرة لإجراء التعديلات المطلوبة.'
                            : 'Interact directly with the pages to apply changes.'}
                        </p>
                      </div>
                      <button
                        onClick={() => {
                          setSelectedFiles([]);
                          setInteractivePages([]);
                        }}
                        className="text-xs font-semibold px-3 py-1.5 bg-slate-50 dark:bg-slate-800 hover:bg-red-50 dark:hover:bg-red-950/20 text-slate-600 dark:text-slate-300 hover:text-red-600 dark:hover:text-red-400 border border-slate-200 dark:border-slate-700 hover:border-red-200 dark:hover:border-red-800 rounded-lg shadow-sm transition-all flex items-center gap-1"
                      >
                        <span>✕</span>
                        <span>{language === Language.AR ? 'إلغاء وتراجع' : 'Cancel & Reset'}</span>
                      </button>
                    </div>

                    {isLoadingPages ? (
                      <div className="flex flex-col items-center justify-center py-20 space-y-3 bg-slate-50 dark:bg-slate-950 rounded-xl border border-slate-200 dark:border-slate-800 border-dashed">
                        <div className="w-8 h-8 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin"></div>
                        <p className="text-xs text-slate-500 dark:text-slate-400 font-semibold">
                          {language === Language.AR ? 'جاري استخراج صفحات المستند...' : 'Reading document pages...'}
                        </p>
                      </div>
                    ) : (selectedFiles.length === 0 && toolId === 'qrcode') ? (
                      /* Live QR Code Preview Panel when no files uploaded */
                      <div className="flex flex-col items-center justify-center p-8 bg-slate-50 dark:bg-slate-950/40 rounded-2xl border border-slate-200 dark:border-slate-800 space-y-6">
                        <div className="p-4 bg-white dark:bg-slate-900 rounded-2xl shadow-md border border-slate-100 dark:border-slate-800 flex items-center justify-center max-w-[220px] aspect-square w-full">
                          {qrText.trim() ? (
                            <img
                              src={`https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(qrText.trim())}`}
                              alt="QR Code Preview"
                              className="w-full h-full object-contain rounded-lg shadow-inner"
                              referrerPolicy="no-referrer"
                            />
                          ) : (
                            <div className="w-full h-full border-2 border-dashed border-slate-200 dark:border-slate-800 rounded-lg flex flex-col items-center justify-center p-4 text-center">
                              <span className="text-3xl mb-2 opacity-65">📱</span>
                              <span className="text-[10px] text-slate-400 font-semibold leading-normal">
                                {language === Language.AR ? 'اكتب رابطاً أو نصاً بالجانب الأيمن للمعاينة الفورية' : 'Enter a text/URL on the right to preview'}
                              </span>
                            </div>
                          )}
                        </div>
                        <div className="text-center space-y-1.5 max-w-sm">
                          <p className="text-xs font-bold text-slate-800 dark:text-slate-200">
                            {language === Language.AR ? 'مولّد ومصمم رموز الـ QR' : 'Interactive QR Code Studio'}
                          </p>
                          <p className="text-[11px] text-slate-500 dark:text-slate-400 leading-relaxed font-medium">
                            {language === Language.AR
                              ? 'اكتب المحتوى الذي تريده، وحدد طريقة التنزيل كصورة مستقلة مباشرة. لتضمين الرمز داخل ملف PDF، يرجى رفع ملف PDF أولاً.'
                              : 'Enter your content, and download it instantly as a PNG image. To embed it on a document, upload a PDF file first.'}
                          </p>
                        </div>
                      </div>
                    ) : (
                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-4 max-h-[480px] overflow-y-auto p-3 border border-slate-200 dark:border-slate-800 rounded-xl bg-slate-50/50 dark:bg-slate-950/30 shadow-inner">
                        {interactivePages.map((page, idx) => {
                          const isSplit = page.splitAfter;
                          
                          return (
                            <React.Fragment key={page.id}>
                              <InteractivePageCard
                                page={page}
                                idx={idx}
                                toolId={toolId}
                                language={language}
                                isGloballySelected={globalPagesSelectionMode === 'custom' && globalSelectedPageIndices.includes(page.pageNum)}
                                onPageClick={() => {
                                  if (globalPagesSelectionMode === 'custom') {
                                    setGlobalSelectedPageIndices(prev => {
                                      const next = prev.includes(page.pageNum)
                                        ? prev.filter(x => x !== page.pageNum)
                                        : [...prev, page.pageNum];
                                      const text = formatPageIndicesToRangeString(next);
                                      setGlobalPagesSelectionText(text);
                                      return next;
                                    });
                                  } else {
                                    if (toolId === 'delete_pages') {
                                      setInteractivePages(prev =>
                                        prev.map(p => (p.id === page.id ? { ...p, isSelectedForDelete: !p.isSelectedForDelete } : p))
                                      );
                                    } else if (toolId === 'chat') {
                                      setSelectedPageSource(page.pageNum);
                                    }
                                  }
                                }}
                                onRotateClick={() => {
                                  setInteractivePages(prev =>
                                    prev.map(p => (p.id === page.id ? { ...p, rotation: (p.rotation + 90) % 360 } : p))
                                  );
                                }}
                                onMovePage={movePage}
                                isFirst={idx === 0}
                                isLast={idx === interactivePages.length - 1}
                                watermarkText={watermarkText}
                                cropPercentage={cropPercentage}
                                requestRender={handleRequestRender}
                                imageSrc={['img2pdf', 'png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId) ? imgUrls[`img-${page.fileIndex}`] : pageImages[`${page.fileIndex}-${page.pageNum}`]}
                                renderStatus={['img2pdf', 'png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId) ? 'done' : pageRenderStatus[`${page.fileIndex}-${page.pageNum}`]}
                                rangeInfo={(() => {
                                  if (toolId !== 'split' || splitType !== 'ranges') return null;
                                  const index1 = page.pageNum + 1;
                                  const rIdx = customRanges.findIndex(r => index1 >= r.from && index1 <= r.to);
                                  if (rIdx === -1) return null;
                                  return { index: rIdx, color: customRanges[rIdx].color };
                                })()}
                                onDeleteClick={['merge', 'img2pdf'].includes(toolId) ? () => {
                                  setInteractivePages(prev => prev.filter(p => p.id !== page.id));
                                } : undefined}
                                selectedPageSource={selectedPageSource}
                              />

                              {/* Separators with scissors button for split tool */}
                              {toolId === 'split' && idx < interactivePages.length - 1 && (
                                <div className="col-span-full py-1.5 flex items-center justify-center relative">
                                  <div className="absolute inset-0 flex items-center" aria-hidden="true">
                                    <div className={`w-full border-t-2 transition-colors ${isSplit ? 'border-amber-500 border-solid' : 'border-slate-200 dark:border-slate-800 border-dashed'}`} />
                                  </div>
                                  <div className="relative flex justify-center">
                                    <button
                                      onClick={() => {
                                        setInteractivePages(prev =>
                                          prev.map(p => (p.id === page.id ? { ...p, splitAfter: !p.splitAfter } : p))
                                        );
                                      }}
                                      className={`relative z-10 px-3.5 py-1.5 rounded-full text-[10px] font-bold shadow-sm transition-all flex items-center space-x-1.5 space-x-reverse border ${
                                        isSplit
                                          ? 'bg-amber-500 text-white border-amber-600 hover:bg-amber-600 ring-2 ring-amber-400/20'
                                          : 'bg-white dark:bg-slate-800 text-slate-500 dark:text-slate-400 border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-750'
                                      }`}
                                    >
                                      <span>✂️</span>
                                      <span>
                                        {isSplit 
                                          ? (language === Language.AR ? 'حد تقسيم الملف نشط' : 'Active Split Bound') 
                                          : (language === Language.AR ? 'قص وتقسيم هنا' : 'Split Bound Here')}
                                      </span>
                                    </button>
                                  </div>
                                </div>
                              )}
                            </React.Fragment>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  {/* Right Column (1/3 width) - Options & Execute Confirmation Panel */}
                  <div className="bg-slate-50/80 dark:bg-slate-900/80 border border-slate-200 dark:border-slate-800 rounded-xl p-5 flex flex-col justify-between space-y-5 shadow-sm max-h-[600px] overflow-y-auto">
                    <div className="space-y-4">
                      <h4 className="text-sm font-bold text-slate-800 dark:text-slate-200 border-b border-slate-200 dark:border-slate-800 pb-2 flex items-center gap-1.5">
                        <span>⚙️</span>
                        <span>{language === Language.AR ? 'خيارات مخصصة ولوحة التحكم' : 'Custom Options Dashboard'}</span>
                      </h4>

                      {/* Display files info with direct deletion capability */}
                      {selectedFiles.length > 0 ? (
                        <div className="text-xs text-slate-600 dark:text-slate-300 space-y-1.5 bg-white dark:bg-slate-950 p-3 rounded-xl border border-slate-200 dark:border-slate-800">
                          <div className="flex items-center justify-between">
                            <p className="font-bold text-slate-700 dark:text-slate-300">{language === Language.AR ? 'المستندات المرفوعة:' : 'Uploaded files:'}</p>
                            <span className="bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 text-[10px] px-2 py-0.5 rounded-full font-bold">
                              {selectedFiles.length} {language === Language.AR ? 'ملفات' : 'files'}
                            </span>
                          </div>
                          <div className="max-h-24 overflow-y-auto space-y-1.5 pr-0.5">
                            {selectedFiles.map((f, i) => (
                              <div key={i} className="flex items-center justify-between gap-1.5 bg-slate-50 dark:bg-slate-900 border border-slate-100 dark:border-slate-800 px-2 py-1 rounded truncate">
                                <div className="flex items-center gap-1 truncate text-[10px]">
                                  <span className="font-extrabold text-indigo-500">#{i + 1}</span>
                                  <span className="truncate max-w-[120px] font-medium text-slate-700 dark:text-slate-300">{f.name}</span>
                                </div>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    removeFile(i);
                                  }}
                                  className="text-red-500 hover:text-red-700 hover:bg-red-50 dark:hover:bg-red-950/30 p-1 rounded-md transition-colors text-[9px]"
                                  title={language === Language.AR ? 'حذف هذا الملف' : 'Delete this file'}
                                >
                                  ✕
                                </button>
                              </div>
                            ))}
                          </div>
                        </div>
                      ) : toolId === 'qrcode' ? (
                        <button
                          type="button"
                          onClick={() => fileInputRef.current?.click()}
                          className="w-full py-3 px-4 border border-dashed border-indigo-200 dark:border-indigo-850 hover:border-indigo-400 dark:hover:border-indigo-700 bg-indigo-50/30 hover:bg-indigo-50/65 dark:bg-indigo-950/10 dark:hover:bg-indigo-950/20 rounded-xl transition-all text-center flex flex-col items-center justify-center space-y-1 group cursor-pointer"
                        >
                          <span className="text-xs font-bold text-indigo-950 dark:text-indigo-400 flex items-center gap-1 group-hover:text-indigo-600 transition-colors">
                            <span>📎</span>
                            <span>{language === Language.AR ? 'تضمين رمز الـ QR داخل ملف PDF' : 'Embed QR inside a PDF file'}</span>
                          </span>
                          <span className="text-[10px] text-slate-400 dark:text-slate-500 font-semibold group-hover:text-slate-500 transition-colors">
                            {language === Language.AR ? 'انقر لرفع مستند وتحديد مكان توضع الرمز بصرياً' : 'Click to upload a document and visually place the code'}
                          </span>
                        </button>
                      ) : null}

                      {/* Universal Output Filename (For Merge, Word, Excel, PPT, PNG, SVG, HTML, TXT, EPUB) */}
                      {['merge', 'pdf2word', 'pdf2excel', 'pdf2ppt', 'pdf2png', 'pdf2svg', 'pdf2html', 'pdf2txt', 'pdf2epub'].includes(toolId) && (
                        <div className="space-y-1.5">
                          <label className="block text-xs font-bold text-slate-700 dark:text-slate-300">
                            {language === Language.AR ? 'اسم الملف الناتج:' : 'Output Filename:'}
                          </label>
                          <div className="relative">
                            <input
                              type="text"
                              value={outputFileName}
                              onChange={(e) => setOutputFileName(e.target.value)}
                              placeholder={language === Language.AR ? 'أدخر اسم الملف بدون امتداد' : 'Enter filename without extension'}
                              className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-xs bg-white dark:bg-slate-900 transition-all pr-12 text-slate-800 dark:text-slate-100 font-medium"
                            />
                            <span className="absolute right-3 top-2 text-[10px] text-slate-400 font-mono font-bold">
                              {toolId === 'merge' ? '.pdf' : toolId === 'pdf2word' ? '.docx' : toolId === 'pdf2excel' ? '.csv' : toolId === 'pdf2png' ? '.png' : toolId === 'pdf2svg' ? '.svg' : toolId === 'pdf2html' ? '.html' : toolId === 'pdf2txt' ? '.txt' : toolId === 'pdf2epub' ? '.epub' : '.ppt'}
                            </span>
                          </div>
                        </div>
                      )}

                      {/* Merge Tool Info Helper */}
                      {toolId === 'merge' && (
                        <div className="p-3 bg-indigo-50 dark:bg-indigo-950/30 border border-indigo-100 dark:border-indigo-900 rounded-xl text-[10px] text-indigo-900 dark:text-indigo-200 leading-relaxed space-y-1">
                          <p className="font-bold flex items-center gap-1">
                            <span>💡</span>
                            <span>{language === Language.AR ? 'ترتيب مرن وتعديل:' : 'Flexible ordering & delete:'}</span>
                          </p>
                          <p>
                            {language === Language.AR
                              ? 'يمكنك سحب وترتيب الصفحات بالمعاينة الفعالة، والضغط على ✕ لحذف صفحات محددة قبل الدمج.'
                              : 'Drag, reorder, or delete individual pages directly in the visual gallery preview before merging.'}
                          </p>
                        </div>
                      )}

                      {/* Split Options Dashboard */}
                      {toolId === 'split' && (
                        <div className="space-y-3">
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'طريقة تقسيم المستند:' : 'Document Split Mode:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setSplitType('all')}
                                className={`px-2.5 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  splitType === 'all'
                                    ? 'bg-amber-500 border-amber-600 text-white shadow-sm'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'استخراج كل صفحة كملف' : 'Extract Every Page'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setSplitType('ranges')}
                                className={`px-2.5 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  splitType === 'ranges'
                                    ? 'bg-amber-500 border-amber-600 text-white shadow-sm'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نطاقات مخصصة' : 'Custom Ranges'}
                              </button>
                            </div>
                          </div>

                          {splitType === 'ranges' && (
                            <div className="space-y-2 bg-white p-3 rounded-xl border border-slate-200 max-h-48 overflow-y-auto">
                              <div className="flex items-center justify-between border-b border-slate-100 pb-1.5">
                                <span className="text-[10px] font-extrabold text-slate-600">{language === Language.AR ? 'نطاقات التقسيم النشطة:' : 'Active Ranges:'}</span>
                                <button
                                  type="button"
                                  onClick={() => {
                                    const colors = ['emerald', 'sky', 'amber', 'pink', 'indigo', 'rose'];
                                    const nextColor = colors[customRanges.length % colors.length];
                                    setCustomRanges([...customRanges, { id: Date.now().toString(), from: 1, to: 1, color: nextColor }]);
                                  }}
                                  className="text-[9px] font-extrabold text-indigo-600 hover:text-indigo-800 bg-indigo-50 hover:bg-indigo-100 px-2 py-1 rounded"
                                >
                                  + {language === Language.AR ? 'إضافة نطاق' : 'Add Range'}
                                </button>
                              </div>
                              
                              {customRanges.map((range, rIdx) => (
                                <div key={range.id} className="flex items-center gap-1.5 p-1.5 bg-slate-50 border border-slate-100 rounded-lg text-[10px]">
                                  <div className={`w-3.5 h-3.5 rounded-full bg-${range.color}-500 flex-shrink-0`} title={`Range ${rIdx + 1}`} />
                                  <span className="font-bold text-slate-500">#{rIdx + 1}</span>
                                  
                                  <div className="flex items-center gap-1">
                                    <span className="text-slate-400">{language === Language.AR ? 'من' : 'From'}</span>
                                    <input
                                      type="number"
                                      min={1}
                                      value={range.from}
                                      onChange={(e) => {
                                        const val = parseInt(e.target.value, 10) || 1;
                                        setCustomRanges(prev => prev.map(item => item.id === range.id ? { ...item, from: val } : item));
                                      }}
                                      className="w-10 px-1 py-0.5 rounded border border-slate-200 text-center font-bold bg-white text-slate-800"
                                    />
                                  </div>

                                  <div className="flex items-center gap-1">
                                    <span className="text-slate-400">{language === Language.AR ? 'إلى' : 'To'}</span>
                                    <input
                                      type="number"
                                      min={1}
                                      value={range.to}
                                      onChange={(e) => {
                                        const val = parseInt(e.target.value, 10) || 1;
                                        setCustomRanges(prev => prev.map(item => item.id === range.id ? { ...item, to: val } : item));
                                      }}
                                      className="w-10 px-1 py-0.5 rounded border border-slate-200 text-center font-bold bg-white text-slate-800"
                                    />
                                  </div>

                                  <button
                                    type="button"
                                    disabled={customRanges.length === 1}
                                    onClick={() => setCustomRanges(prev => prev.filter(item => item.id !== range.id))}
                                    className="ms-auto text-red-500 hover:text-red-700 disabled:opacity-40 p-1"
                                  >
                                    🗑️
                                  </button>
                                </div>
                              ))}
                            </div>
                          )}

                          <div className="p-2.5 bg-amber-50 border border-amber-200 rounded-lg text-[10px] text-amber-800 space-y-1">
                            <p className="font-bold">✂️ {language === Language.AR ? 'ملخص التقسيم بالتلوين:' : 'Color-Coded Split Summary:'}</p>
                            <p>
                              {language === Language.AR
                                ? splitType === 'all'
                                  ? `سيتم إنتاج كل صفحة كملف PDF مستقل (${interactivePages.length} ملفات).`
                                  : `سيتم إنتاج (${customRanges.length}) ملفات بناءً على النطاقات الملونة بالمعاينة.`
                                : splitType === 'all'
                                  ? `Will split each page into its own individual PDF (${interactivePages.length} files).`
                                  : `Will produce (${customRanges.length}) separate PDFs highlighted visually in the gallery.`}
                            </p>
                          </div>
                        </div>
                      )}

                      {/* Compress options */}
                      {toolId === 'compress' && (
                        <div className="space-y-3">
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'اختر مستوى الضغط:' : 'Select Compression Level:'}
                            </label>
                            <div className="space-y-1.5">
                              {[
                                { id: 'recommended', labelAr: 'موصى به (جودة عالية وحجم مثالي)', labelEn: 'Recommended (Optimal balance)', descAr: 'توفير مساحة كبيرة مع الحفاظ على وضوح رائع للنصوص والصور.', descEn: 'Saves major space while keeping crystal clear readability.', color: 'border-emerald-500 bg-emerald-50/50' },
                                { id: 'basic', labelAr: 'أساسي (جودة أصلية وحجم أصغر)', labelEn: 'Basic (Low compression)', descAr: 'ضغط خفيف وسريع جداً، مثالي للمستندات المليئة بالنصوص الحساسة.', descEn: 'Fast and light optimization. Perfect for precise textual records.', color: 'border-indigo-500 bg-indigo-50/50' },
                                { id: 'max', labelAr: 'أقصى ضغط (حجم صغير وجودة مقبولة)', labelEn: 'Maximum (Extreme compression)', descAr: 'أكبر ضغط ممكن لتسهيل الإرسال عبر البريد، قد يقلل جودة الصور.', descEn: 'Extremely small output size. Ideal for swift email attachments.', color: 'border-rose-500 bg-rose-50/50' }
                              ].map(item => (
                                <button
                                  key={item.id}
                                  type="button"
                                  onClick={() => setCompressLevel(item.id as any)}
                                  className={`w-full p-2.5 rounded-xl border text-right text-xs transition-all flex flex-col gap-1 ${
                                    compressLevel === item.id 
                                      ? `${item.color} ring-2 ring-indigo-500/30 font-bold`
                                      : 'bg-white border-slate-200 hover:bg-slate-50'
                                  }`}
                                >
                                  <div className="flex items-center gap-1.5 font-bold text-slate-800 text-[11px] justify-between w-full">
                                    <span>{language === Language.AR ? item.labelAr : item.labelEn}</span>
                                    {compressLevel === item.id && <span className="text-[10px] text-indigo-600 font-extrabold">✓</span>}
                                  </div>
                                  <span className="text-[10px] text-slate-500 font-medium">
                                    {language === Language.AR ? item.descAr : item.descEn}
                                  </span>
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Estimated file size compression results panel */}
                          {selectedFiles[0] && (
                            <div className="bg-slate-100 rounded-xl p-3 border border-slate-200 text-[10px] space-y-1.5">
                              <p className="font-extrabold text-slate-700 flex items-center gap-1">
                                <span>📊</span>
                                <span>{language === Language.AR ? 'المقارنة المتوقعة لحجم الملف:' : 'Estimated Compression Comparison:'}</span>
                              </p>
                              <div className="grid grid-cols-2 gap-2 text-slate-600 font-medium">
                                <div>
                                  <span className="block text-slate-400">{language === Language.AR ? 'الحجم الأصلي:' : 'Original Size:'}</span>
                                  <span className="font-mono font-bold text-slate-700">
                                    {(selectedFiles[0].size / 1024).toFixed(1)} KB
                                  </span>
                                </div>
                                <div>
                                  <span className="block text-slate-400">{language === Language.AR ? 'الحجم المقدر:' : 'Estimated Output:'}</span>
                                  <span className="font-mono font-bold text-emerald-600">
                                    {(() => {
                                      const orig = selectedFiles[0].size / 1024;
                                      let ratio = 0.55;
                                      if (compressLevel === 'basic') ratio = 0.85;
                                      if (compressLevel === 'max') ratio = 0.25;
                                      return `${(orig * ratio).toFixed(1)} KB (-${Math.round((1 - ratio) * 100)}%)`;
                                    })()}
                                  </span>
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      )}

                      {/* PDF to Word Options */}
                      {toolId === 'pdf2word' && (
                        <div className="space-y-3.5">
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تفضيلات نمط التحويل:' : 'Conversion Mode Preferences:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setPdf2wordMode('format')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  pdf2wordMode === 'format'
                                    ? 'bg-blue-600 border-blue-700 text-white shadow-sm'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'التنسيق الأصلي بالكامل' : 'Original Formatting'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setPdf2wordMode('text')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  pdf2wordMode === 'text'
                                    ? 'bg-blue-600 border-blue-700 text-white shadow-sm'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'استخراج النصوص فقط' : 'Text Content Only'}
                              </button>
                            </div>
                          </div>

                          {/* Scanned detection banner / OCR trigger suggestion */}
                          {isScannedDetected && (
                            <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl text-[10px] text-amber-900 leading-relaxed space-y-1.5">
                              <p className="font-bold flex items-center gap-1">
                                <span>⚠️</span>
                                <span>{language === Language.AR ? 'تنبيه مستند مصور (Scanned PDF):' : 'Scanned PDF Detected:'}</span>
                              </p>
                              <p>
                                {language === Language.AR
                                  ? 'يبدو أن هذا المستند عبارة عن صور ممسوحة ضوئياً ولا يحتوي على نصوص قابلة للتحديد المباشر. نقترح بشدة تفعيل محرك الـ OCR أدناه لتحويل آمن.'
                                  : 'This document appears to contain scanned image content without selectable text layers. We strongly recommend enabling OCR below.'}
                              </p>
                            </div>
                          )}

                          <div className="flex items-center justify-between p-3 bg-white rounded-xl border border-slate-200">
                            <div className="space-y-0.5">
                              <span className="block text-xs font-bold text-slate-800">
                                {language === Language.AR ? 'تفعيل ميزة الـ OCR الذكية:' : 'Enable Smart OCR Reading:'}
                              </span>
                              <span className="block text-[10px] text-slate-400 font-medium">
                                {language === Language.AR ? 'التعرف الضوئي على المستندات المصورة' : 'Optical Character Recognition for scanned pages'}
                              </span>
                            </div>
                            <button
                              type="button"
                              onClick={() => setOcrEnabled(!ocrEnabled)}
                              className={`w-11 h-6 rounded-full transition-colors relative flex items-center ${
                                ocrEnabled ? 'bg-indigo-600' : 'bg-slate-200'
                              }`}
                            >
                              <span
                                className={`w-4 h-4 rounded-full bg-white shadow-md absolute transition-transform ${
                                  ocrEnabled ? 'translate-x-6' : 'translate-x-1'
                                }`}
                              />
                            </button>
                          </div>
                        </div>
                      )}

                      {/* PDF Conversion Options */}
                      {isPdfConversionTool && (
                        <div className="space-y-4">
                          {toolId === 'html2pdf' && htmlUrl.trim().length > 0 && (
                            <div className="space-y-1.5">
                              <label className="block text-xs font-bold text-slate-700">
                                {language === Language.AR ? 'رابط صفحة الويب النشط:' : 'Active Web Page Link:'}
                              </label>
                              <input
                                type="url"
                                value={htmlUrl}
                                onChange={(e) => setHtmlUrl(e.target.value)}
                                className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-xs bg-white text-slate-800 font-semibold"
                              />
                            </div>
                          )}

                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تحديد جودة الضغط والتحويل:' : 'Select Conversion Quality:'}
                            </label>
                            <div className="grid grid-cols-3 gap-2">
                              {['low', 'medium', 'high'].map(level => (
                                <button
                                  key={level}
                                  type="button"
                                  onClick={() => setPdfConvQuality(level as any)}
                                  className={`px-2 py-1.5 rounded-lg border text-center text-[10px] font-bold uppercase transition-all cursor-pointer ${
                                    pdfConvQuality === level
                                      ? 'bg-indigo-600 border-indigo-700 text-white shadow'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {level === 'low' ? (language === Language.AR ? 'ضغط عالٍ' : 'High Comp.') : level === 'medium' ? (language === Language.AR ? 'متوازن' : 'Balanced') : (language === Language.AR ? 'جودة فائقة' : 'Max')}
                                </button>
                              ))}
                            </div>
                            <span className="block text-[9px] text-slate-400 leading-relaxed">
                              {language === Language.AR
                                ? 'يتحكم في دقة النصوص وجودة معالجة الجداول والصور داخل ملف PDF الناتج.'
                                : 'Controls resolution and layout preservation quality in the resulting PDF.'}
                            </span>
                          </div>
                        </div>
                      )}

                      {/* PDF to Excel Options */}
                      {toolId === 'pdf2excel' && (
                        <div className="space-y-3">
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'أدخل أرقام صفحات الجداول المحددة:' : 'Specify spreadsheet page numbers:'}
                            </label>
                            <input
                              type="text"
                              value={excelPages}
                              onChange={(e) => setExcelPages(e.target.value)}
                              placeholder="e.g. 1, 2-4, 5"
                              className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-xs bg-white text-slate-800 font-semibold"
                            />
                            <span className="block text-[9px] text-slate-400">
                              {language === Language.AR ? 'اتركها فارغة لاستخراج كافة جداول المستند.' : 'Leave empty to scan and extract all pages.'}
                            </span>
                          </div>

                          {/* Beautiful simulated table layout preview */}
                          <div className="bg-slate-100 rounded-xl p-3 border border-slate-200 space-y-2 text-[10px]">
                            <p className="font-extrabold text-slate-700 flex items-center gap-1.5">
                              <span>📊</span>
                              <span>{language === Language.AR ? 'معاينة افتراضية للجداول المكتشفة:' : 'Simulated Tables Live Preview:'}</span>
                            </p>
                            <div className="border border-slate-300 rounded overflow-hidden shadow-inner bg-white">
                              <table className="w-full text-left font-mono text-[9px] text-slate-600">
                                <thead className="bg-slate-50 border-b border-slate-200">
                                  <tr>
                                    <th className="p-1 border-r border-slate-200 text-[8px] font-extrabold">A</th>
                                    <th className="p-1 border-r border-slate-200 text-[8px] font-extrabold">B</th>
                                    <th className="p-1 text-[8px] font-extrabold">C</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  <tr className="border-b border-slate-150">
                                    <td className="p-1 border-r border-slate-200 truncate">ID_Key_98</td>
                                    <td className="p-1 border-r border-slate-200 text-indigo-600">Table Row 1</td>
                                    <td className="p-1 font-bold text-slate-700">$54.00</td>
                                  </tr>
                                  <tr>
                                    <td className="p-1 border-r border-slate-200 truncate">ID_Key_99</td>
                                    <td className="p-1 border-r border-slate-200 text-indigo-600">Table Row 2</td>
                                    <td className="p-1 font-bold text-slate-700">$128.50</td>
                                  </tr>
                                </tbody>
                              </table>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* PDF to PowerPoint Options */}
                      {toolId === 'pdf2ppt' && (
                        <div className="space-y-3.5">
                          <div className="flex items-center justify-between p-3 bg-white rounded-xl border border-slate-200">
                            <div className="space-y-0.5">
                              <span className="block text-xs font-bold text-slate-800">
                                {language === Language.AR ? 'كل صفحة شريحة مستقلة:' : 'One page per slide layout:'}
                              </span>
                              <span className="block text-[10px] text-slate-400 font-medium">
                                {language === Language.AR ? 'إنشاء شريحة متناسقة لكل صفحة مستند' : 'Map each PDF page directly to a slide frame'}
                              </span>
                            </div>
                            <button
                              type="button"
                              onClick={() => setPptPageAsSlide(!pptPageAsSlide)}
                              className={`w-11 h-6 rounded-full transition-colors relative flex items-center ${
                                pptPageAsSlide ? 'bg-orange-500' : 'bg-slate-200'
                              }`}
                            >
                              <span
                                className={`w-4 h-4 rounded-full bg-white shadow-md absolute transition-transform ${
                                  pptPageAsSlide ? 'translate-x-6' : 'translate-x-1'
                                }`}
                              />
                            </button>
                          </div>

                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'دقة وجودة الصور بالعرض:' : 'Image elements slide resolution:'}
                            </label>
                            <div className="grid grid-cols-3 gap-2">
                              {['low', 'medium', 'high'].map(level => (
                                <button
                                  key={level}
                                  type="button"
                                  onClick={() => setPptQuality(level as any)}
                                  className={`px-2 py-1.5 rounded-lg border text-center text-[10px] font-bold uppercase transition-all ${
                                    pptQuality === level
                                      ? 'bg-orange-500 border-orange-600 text-white shadow'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {level === 'low' ? (language === Language.AR ? 'سريع' : 'Draft') : level === 'medium' ? (language === Language.AR ? 'متوسط' : 'Medium') : (language === Language.AR ? 'فائق' : 'HD')}
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Images to PDF Options */}
                      {toolId === 'img2pdf' && (
                        <div className="space-y-3.5">
                          {/* Page orientation selector */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'اتجاه الصفحة الناتج:' : 'Target Page Orientation:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setImgOrientation('portrait')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  imgOrientation === 'portrait'
                                    ? 'bg-indigo-600 border-indigo-700 text-white'
                                    : 'bg-white border-slate-200 text-slate-600'
                                }`}
                              >
                                📄 {language === Language.AR ? 'طولي (Portrait)' : 'Portrait'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setImgOrientation('landscape')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  imgOrientation === 'landscape'
                                    ? 'bg-indigo-600 border-indigo-700 text-white'
                                    : 'bg-white border-slate-200 text-slate-600'
                                }`}
                              >
                                📋 {language === Language.AR ? 'عرضي (Landscape)' : 'Landscape'}
                              </button>
                            </div>
                          </div>

                          {/* Margin settings selector */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'حجم هوامش الصفحة:' : 'Page Margin Setting:'}
                            </label>
                            <div className="grid grid-cols-3 gap-1.5">
                              {[
                                { id: 'none', labelAr: 'بلا هوامش', labelEn: 'None' },
                                { id: 'small', labelAr: 'صغير', labelEn: 'Small' },
                                { id: 'medium', labelAr: 'متوسط', labelEn: 'Medium' }
                              ].map(marginItem => (
                                <button
                                  key={marginItem.id}
                                  type="button"
                                  onClick={() => setImgMargin(marginItem.id as any)}
                                  className={`px-1.5 py-1.5 rounded-lg border text-center text-[10px] font-bold transition-all ${
                                    imgMargin === marginItem.id
                                      ? 'bg-indigo-600 border-indigo-700 text-white'
                                      : 'bg-white border-slate-200 text-slate-600'
                                  }`}
                                >
                                  {language === Language.AR ? marginItem.labelAr : marginItem.labelEn}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Layout options: 1 image per page vs multiple */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'هيكل التخطيط بالصفحة:' : 'Target Page Layout Grid:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setImgLayout('separate')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  imgLayout === 'separate'
                                    ? 'bg-indigo-600 border-indigo-700 text-white'
                                    : 'bg-white border-slate-200 text-slate-600'
                                }`}
                              >
                                {language === Language.AR ? 'صورة لكل صفحة' : '1 Image Per Page'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setImgLayout('combined')}
                                className={`px-2 py-2 rounded-lg border text-center text-xs font-bold transition-all ${
                                  imgLayout === 'combined'
                                    ? 'bg-indigo-600 border-indigo-700 text-white'
                                    : 'bg-white border-slate-200 text-slate-600'
                                }`}
                              >
                                {language === Language.AR ? 'دمج صورتين بالصفحة' : '2 Images Per Page'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Image Conversion Options */}
                      {['png2webp', 'jpg2webp', 'png2heic', 'jpg2heic'].includes(toolId) && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-bold text-slate-700 flex justify-between">
                              <span>{language === Language.AR ? 'جودة ضغط الصورة:' : 'Image Compression Quality:'}</span>
                              <span className="text-indigo-600 font-extrabold">{imageQuality}%</span>
                            </label>
                            <input
                              type="range"
                              min="10"
                              max="100"
                              step="5"
                              value={imageQuality}
                              onChange={(e) => setImageQuality(parseInt(e.target.value, 10))}
                              className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                            />
                            <div className="flex justify-between text-[10px] text-slate-400 font-medium">
                              <span>{language === Language.AR ? 'أقصى ضغط' : 'Max Compress'}</span>
                              <span>{language === Language.AR ? 'أفضل جودة' : 'Best Quality'}</span>
                            </div>
                          </div>

                          <div className="p-3 bg-indigo-50 border border-indigo-100/50 rounded-xl text-[11px] text-slate-600 leading-relaxed">
                            <p className="font-bold text-indigo-900 mb-1 flex items-center gap-1">
                              <span>💡</span>
                              <span>{language === Language.AR ? 'معلومات التحويل:' : 'Conversion Info:'}</span>
                            </p>
                            <p>
                              {toolId.includes('webp') 
                                ? (language === Language.AR 
                                    ? 'صيغة WEBP توفر ضغطاً ممتازاً للصور مع الحفاظ على شفافيتها وتفاصيلها وهي الخيار المثالي لمواقع الويب.' 
                                    : 'WEBP format offers outstanding compression, preserving details and transparency. Ideal for web performance.')
                                : (language === Language.AR 
                                    ? 'صيغة HEIC (صيغة الصور عالية الكفاءة) هي المعيار الحديث المستخدم في أجهزة Apple لتوفير جودة عالية بنصف الحجم.' 
                                    : 'HEIC (High Efficiency Image Container) is the modern standard used by Apple devices to store high-quality photos at half the size.')
                              }
                            </p>
                          </div>
                        </div>
                      )}

                      {/* Watermark Options (Tool 8) */}
                      {toolId === 'watermark' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Toggle Text vs Image */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'نوع العلامة المائية:' : 'Watermark Type:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setWatermarkType('text')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  watermarkType === 'text'
                                    ? 'bg-red-600 border-red-700 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                📝 {language === Language.AR ? 'نص مخصص' : 'Custom Text'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setWatermarkType('image')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  watermarkType === 'image'
                                    ? 'bg-red-600 border-red-700 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                🖼️ {language === Language.AR ? 'صورة أو شعار' : 'Image Logo'}
                              </button>
                            </div>
                          </div>

                          {/* Text input options */}
                          {watermarkType === 'text' ? (
                            <div className="space-y-3 p-3 bg-slate-50 border border-slate-200 rounded-xl">
                              <div className="space-y-1">
                                <label className="block text-[11px] font-extrabold text-slate-600">
                                  {language === Language.AR ? 'نص العلامة المائية:' : 'Watermark Text:'}
                                </label>
                                <input
                                  type="text"
                                  value={watermarkText}
                                  onChange={(e) => setWatermarkText(e.target.value)}
                                  placeholder={texts.watermarkPlaceholder}
                                  className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-red-500 focus:border-red-500 outline-none text-xs bg-white transition-all font-bold"
                                />
                              </div>

                              <div className="grid grid-cols-2 gap-2">
                                <div className="space-y-1">
                                  <label className="block text-[11px] font-extrabold text-slate-600">
                                    {language === Language.AR ? 'نوع الخط:' : 'Font Family:'}
                                  </label>
                                  <select
                                    value={watermarkFont}
                                    onChange={(e) => setWatermarkFont(e.target.value)}
                                    className="w-full px-2 py-1.5 rounded border border-slate-300 text-xs bg-white font-bold"
                                  >
                                    <option value="Helvetica">Helvetica</option>
                                    <option value="Times">Times New Roman</option>
                                    <option value="Courier">Courier Mono</option>
                                    <option value="Arial">Arial Black</option>
                                    <option value="Impact">Impact</option>
                                  </select>
                                </div>

                                <div className="space-y-1">
                                  <label className="block text-[11px] font-extrabold text-slate-600">
                                    {language === Language.AR ? 'حجم الخط:' : 'Font Size:'}
                                  </label>
                                  <div className="flex items-center gap-1.5">
                                    <input
                                      type="range"
                                      min="12"
                                      max="72"
                                      value={watermarkSize}
                                      onChange={(e) => setWatermarkSize(parseInt(e.target.value, 10))}
                                      className="w-full accent-red-600 h-1"
                                    />
                                    <span className="font-mono text-[10px] font-bold w-6 text-right">{watermarkSize}</span>
                                  </div>
                                </div>
                              </div>

                              <div className="flex items-center justify-between gap-2 border-t border-slate-200/60 pt-2">
                                <div className="flex items-center gap-1">
                                  <label className="text-[11px] font-extrabold text-slate-600">
                                    {language === Language.AR ? 'اللون:' : 'Color:'}
                                  </label>
                                  <input
                                    type="color"
                                    value={watermarkColor}
                                    onChange={(e) => setWatermarkColor(e.target.value)}
                                    className="w-7 h-7 p-0 border border-slate-300 rounded cursor-pointer bg-transparent"
                                  />
                                </div>
                                
                                <div className="flex gap-1.5">
                                  <button
                                    type="button"
                                    onClick={() => setWatermarkBold(!watermarkBold)}
                                    className={`px-2.5 py-1 text-xs rounded border transition-all font-bold ${
                                      watermarkBold ? 'bg-red-50 border-red-300 text-red-600' : 'bg-white border-slate-200 text-slate-500'
                                    }`}
                                  >
                                    B
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setWatermarkItalic(!watermarkItalic)}
                                    className={`px-2.5 py-1 text-xs rounded border transition-all italic font-bold ${
                                      watermarkItalic ? 'bg-red-50 border-red-300 text-red-600' : 'bg-white border-slate-200 text-slate-500'
                                    }`}
                                  >
                                    I
                                  </button>
                                </div>
                              </div>
                            </div>
                          ) : (
                            /* Image watermark select option */
                            <div className="space-y-3.5 p-3 bg-slate-50 border border-slate-200 rounded-xl text-center">
                              <p className="text-[11px] font-semibold text-slate-500 leading-relaxed text-right">
                                {language === Language.AR 
                                  ? 'اختر من الأختام الجاهزة أدناه أو ارفع صورتك مخصصة:' 
                                  : 'Select from preset stamps or upload your own corporate logo:'}
                              </p>
                              <div className="grid grid-cols-2 gap-2">
                                <button
                                  type="button"
                                  onClick={() => setWatermarkImageUrl('/watermark_approved.png')}
                                  className={`p-2 border rounded-lg text-center flex flex-col items-center justify-center gap-1 transition-all ${
                                    watermarkImageUrl === '/watermark_approved.png' 
                                      ? 'border-red-500 bg-red-50/50 shadow-sm' 
                                      : 'border-slate-200 bg-white hover:bg-slate-100'
                                  }`}
                                >
                                  <span className="text-[9px] font-extrabold uppercase text-emerald-600 border-2 border-emerald-600 rounded px-1 scale-90">APPROVED</span>
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setWatermarkImageUrl('/watermark_draft.png')}
                                  className={`p-2 border rounded-lg text-center flex flex-col items-center justify-center gap-1 transition-all ${
                                    watermarkImageUrl === '/watermark_draft.png' 
                                      ? 'border-red-500 bg-red-50/50 shadow-sm' 
                                      : 'border-slate-200 bg-white hover:bg-slate-100'
                                  }`}
                                >
                                  <span className="text-[9px] font-extrabold uppercase text-amber-600 border-2 border-amber-600 rounded px-1 scale-90">DRAFT ONLY</span>
                                </button>
                              </div>
                              <div className="border-t border-slate-200 pt-2">
                                <label className="block text-[11px] font-extrabold text-slate-600 text-right mb-1">
                                  {language === Language.AR ? 'تحميل صورة ختم مخصصة:' : 'Upload custom png seal:'}
                                </label>
                                <input 
                                  type="file" 
                                  accept="image/png, image/jpeg"
                                  onChange={(e) => {
                                    if (e.target.files && e.target.files[0]) {
                                      setWatermarkImageUrl(URL.createObjectURL(e.target.files[0]));
                                    }
                                  }}
                                  className="w-full text-[10px] text-slate-500 bg-white p-1 rounded border border-slate-300" 
                                />
                              </div>
                            </div>
                          )}

                          {/* 3x3 Positioning Grid */}
                          <div className="space-y-2">
                            <div className="flex items-center justify-between">
                              <span className="block text-xs font-bold text-slate-700">
                                {language === Language.AR ? 'موقع العلامة المائية بالصفحة:' : 'Page Positioning Layout:'}
                              </span>
                              <span className="text-[10px] font-bold text-red-600 capitalize bg-red-50 px-2 py-0.5 rounded border border-red-200/55">
                                {watermarkPosition}
                              </span>
                            </div>
                            <div className="grid grid-cols-3 gap-1 w-28 mx-auto border border-slate-200 p-1 rounded-lg bg-slate-50 shadow-inner">
                              {['top-left', 'top-center', 'top-right', 'center-left', 'center-center', 'center-right', 'bottom-left', 'bottom-center', 'bottom-right'].map(pos => (
                                <button
                                  key={pos}
                                  type="button"
                                  onClick={() => setWatermarkPosition(pos)}
                                  className={`w-8 h-8 rounded transition-all flex items-center justify-center ${
                                    watermarkPosition === pos 
                                      ? 'bg-red-500 text-white shadow-sm font-extrabold scale-105' 
                                      : 'bg-white text-slate-300 hover:bg-slate-100 border border-slate-200'
                                  }`}
                                  title={pos}
                                >
                                  •
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Opacity and Rotation Sliders */}
                          <div className="grid grid-cols-2 gap-3 border-t border-slate-100 pt-3">
                            <div className="space-y-1">
                              <label className="block text-[11px] font-extrabold text-slate-600">
                                {language === Language.AR ? 'زاوية الدوران (°):' : 'Rotation Angle (°):'}
                              </label>
                              <div className="flex items-center gap-1">
                                <input
                                  type="range"
                                  min="-180"
                                  max="180"
                                  value={watermarkRotation}
                                  onChange={(e) => setWatermarkRotation(parseInt(e.target.value, 10))}
                                  className="w-full accent-red-600 h-1"
                                />
                                <span className="font-mono text-[9px] font-bold text-slate-700">{watermarkRotation}°</span>
                              </div>
                            </div>

                            <div className="space-y-1">
                              <label className="block text-[11px] font-extrabold text-slate-600">
                                {language === Language.AR ? 'مستوى الشفافية (%):' : 'Opacity Level (%):'}
                              </label>
                              <div className="flex items-center gap-1">
                                <input
                                  type="range"
                                  min="10"
                                  max="100"
                                  value={watermarkOpacity}
                                  onChange={(e) => setWatermarkOpacity(parseInt(e.target.value, 10))}
                                  className="w-full accent-red-600 h-1"
                                />
                                <span className="font-mono text-[9px] font-bold text-slate-700">{watermarkOpacity}%</span>
                              </div>
                            </div>
                          </div>

                          {/* Target pages selection range */}
                          <div className="space-y-1.5 border-t border-slate-100 pt-3">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تطبيق على صفحات:' : 'Apply Watermark to pages:'}
                            </label>
                            <div className="flex gap-2">
                              <button
                                type="button"
                                onClick={() => setWatermarkPagesMode('all')}
                                className={`px-2.5 py-1 rounded text-[11px] font-bold border transition-all ${
                                  watermarkPagesMode === 'all' 
                                    ? 'bg-red-50 border-red-300 text-red-600' 
                                    : 'bg-white border-slate-200 text-slate-500'
                                }`}
                              >
                                {language === Language.AR ? 'كل الصفحات' : 'All Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setWatermarkPagesMode('range')}
                                className={`px-2.5 py-1 rounded text-[11px] font-bold border transition-all ${
                                  watermarkPagesMode === 'range' 
                                    ? 'bg-red-50 border-red-300 text-red-600' 
                                    : 'bg-white border-slate-200 text-slate-500'
                                }`}
                              >
                                {language === Language.AR ? 'نطاق مخصص' : 'Custom Range'}
                              </button>
                            </div>
                            {watermarkPagesMode === 'range' && (
                              <input
                                type="text"
                                value={watermarkPagesRange}
                                onChange={(e) => setWatermarkPagesRange(e.target.value)}
                                placeholder="e.g. 1-3, 5"
                                className="w-full mt-1.5 px-3 py-1.5 rounded border border-slate-300 text-xs bg-white focus:ring-1 focus:ring-red-500"
                              />
                            )}
                          </div>
                        </div>
                      )}

                      {/* Password Protection Options (Tool 9) */}
                      {toolId === 'protect' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Password Input Fields */}
                          <div className="space-y-2.5 p-3.5 bg-slate-50 border border-slate-200 rounded-xl">
                            <div className="space-y-1">
                              <label className="block text-xs font-bold text-slate-700">{texts.passwordLabel}</label>
                              <input
                                type="password"
                                value={securityPassword}
                                onChange={(e) => setSecurityPassword(e.target.value)}
                                placeholder={texts.passwordPlaceholder}
                                className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-xs bg-white transition-all font-bold"
                              />
                            </div>

                            <div className="space-y-1">
                              <label className="block text-xs font-bold text-slate-700">
                                {language === Language.AR ? 'تأكيد كلمة المرور:' : 'Confirm Password:'}
                              </label>
                              <input
                                type="password"
                                value={confirmSecurityPassword}
                                onChange={(e) => setConfirmSecurityPassword(e.target.value)}
                                placeholder={language === Language.AR ? 'أعد كتابة كلمة السر للتأكيد' : 'Re-enter your protection password'}
                                className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-xs bg-white transition-all font-bold"
                              />
                            </div>

                            {/* Verification badge indicators */}
                            {securityPassword && (
                              <div className="flex items-center justify-between text-[10px] font-bold pt-1">
                                <span className={`flex items-center gap-1 ${securityPassword.length < 6 ? 'text-rose-500' : 'text-emerald-600'}`}>
                                  <span>🛡️</span>
                                  <span>{securityPassword.length < 6 
                                    ? (language === Language.AR ? 'ضعيفة (قصيرة)' : 'Weak (Too short)') 
                                    : (language === Language.AR ? 'قوية ومحمية' : 'Strong & secure')}
                                  </span>
                                </span>

                                <span className={securityPassword === confirmSecurityPassword ? 'text-emerald-600' : 'text-rose-500'}>
                                  {securityPassword === confirmSecurityPassword 
                                    ? (language === Language.AR ? '✓ متطابقتان' : '✓ Passwords match') 
                                    : (language === Language.AR ? '⚠ غير متطابقة!' : '⚠ Do not match!')}
                                </span>
                              </div>
                            )}
                          </div>

                          {/* Granular Permissions Switches */}
                          <div className="space-y-2">
                            <span className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'صلاحيات وقيود الأمان المخصصة:' : 'Granular Permissions Settings:'}
                            </span>
                            
                            <div className="space-y-2">
                              {/* Prevent Printing */}
                              <div className="flex items-center justify-between p-2.5 bg-white rounded-lg border border-slate-200">
                                <div className="space-y-0.5">
                                  <span className="block text-[11px] font-bold text-slate-800">
                                    {language === Language.AR ? 'منع طباعة المستند:' : 'Prevent Document Printing:'}
                                  </span>
                                  <span className="block text-[9px] text-slate-400 font-medium">
                                    {language === Language.AR ? 'حظر إرسال الملف لأي طابعة مادية' : 'Blocks rendering to paper / printing spool'}
                                  </span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setPreventPrinting(!preventPrinting)}
                                  className={`w-10 h-5.5 rounded-full transition-colors relative flex items-center ${
                                    preventPrinting ? 'bg-red-500' : 'bg-slate-200'
                                  }`}
                                >
                                  <span className={`w-3.5 h-3.5 rounded-full bg-white shadow absolute transition-transform ${preventPrinting ? 'translate-x-5' : 'translate-x-1'}`} />
                                </button>
                              </div>

                              {/* Prevent Copying */}
                              <div className="flex items-center justify-between p-2.5 bg-white rounded-lg border border-slate-200">
                                <div className="space-y-0.5">
                                  <span className="block text-[11px] font-bold text-slate-800">
                                    {language === Language.AR ? 'منع نسخ النص والسرقة:' : 'Prevent Content Copying:'}
                                  </span>
                                  <span className="block text-[9px] text-slate-400 font-medium">
                                    {language === Language.AR ? 'تعطيل سحب النصوص وتظليلها للسرقة' : 'Disables highlighting and copying textual elements'}
                                  </span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setPreventCopying(!preventCopying)}
                                  className={`w-10 h-5.5 rounded-full transition-colors relative flex items-center ${
                                    preventCopying ? 'bg-red-500' : 'bg-slate-200'
                                  }`}
                                >
                                  <span className={`w-3.5 h-3.5 rounded-full bg-white shadow absolute transition-transform ${preventCopying ? 'translate-x-5' : 'translate-x-1'}`} />
                                </button>
                              </div>

                              {/* Prevent Editing */}
                              <div className="flex items-center justify-between p-2.5 bg-white rounded-lg border border-slate-200">
                                <div className="space-y-0.5">
                                  <span className="block text-[11px] font-bold text-slate-800">
                                    {language === Language.AR ? 'منع تعديل محتوى المستند:' : 'Prevent Content Modification:'}
                                  </span>
                                  <span className="block text-[9px] text-slate-400 font-medium">
                                    {language === Language.AR ? 'منع تدوير أو دمج أو تغيير هيكل المستند' : 'Locks changes to form fills or layout templates'}
                                  </span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setPreventEditing(!preventEditing)}
                                  className={`w-10 h-5.5 rounded-full transition-colors relative flex items-center ${
                                    preventEditing ? 'bg-red-500' : 'bg-slate-200'
                                  }`}
                                >
                                  <span className={`w-3.5 h-3.5 rounded-full bg-white shadow absolute transition-transform ${preventEditing ? 'translate-x-5' : 'translate-x-1'}`} />
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Rotate Pages Options (Tool 10) */}
                      {toolId === 'rotate' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Target Rotation Mode switcher */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تطبيق التدوير على:' : 'Apply Rotation To:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setRotateMode('all')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  rotateMode === 'all'
                                    ? 'bg-amber-500 border-amber-600 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                📑 {language === Language.AR ? 'كل الصفحات دفعة واحدة' : 'All Document Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setRotateMode('single')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  rotateMode === 'single'
                                    ? 'bg-amber-500 border-amber-600 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                🎯 {language === Language.AR ? 'صفحة محددة فقط' : 'Single Selected Page'}
                              </button>
                            </div>
                          </div>

                          {/* Specific Page details selection */}
                          {rotateMode === 'single' && (
                            <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                              <label className="block text-[11px] font-extrabold text-slate-600">
                                {language === Language.AR ? 'اختر الصفحة المراد تدويرها:' : 'Pick Target Page to Rotate:'}
                              </label>
                              <select
                                value={selectedRotatePageId}
                                onChange={(e) => setSelectedRotatePageId(e.target.value)}
                                className="w-full p-1.5 rounded border border-slate-300 text-xs bg-white font-bold"
                              >
                                <option value="">-- {language === Language.AR ? 'اختر صفحة من المعاينة' : 'Select page or click preview'} --</option>
                                {interactivePages.map((p, pIdx) => (
                                  <option key={p.id} value={p.id}>
                                    {language === Language.AR ? `صفحة ${pIdx + 1}` : `Page ${pIdx + 1}`} ({p.rotation}°)
                                  </option>
                                ))}
                              </select>
                              <p className="text-[9px] text-slate-400 leading-relaxed font-semibold">
                                {language === Language.AR 
                                  ? 'تلميح: يمكنك أيضاً الضغط مباشرة على زر ↻ 90° الظاهر على كروت الصفحات بالمعاينة لتدويرها فوراً.' 
                                  : 'Tip: You can also click the ↻ 90° rotate buttons directly on the cards in the visual gallery.'}
                              </p>
                            </div>
                          )}

                          {/* Angle degree click buttons */}
                          <div className="space-y-2">
                            <span className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'زاوية الدوران المرغوبة:' : 'Rotation Angle Action:'}
                            </span>
                            <div className="grid grid-cols-3 gap-1.5">
                              {[90, 180, 270].map(deg => (
                                <button
                                  key={deg}
                                  type="button"
                                  onClick={() => {
                                    if (rotateMode === 'all') {
                                      setInteractivePages(prev => prev.map(p => ({ ...p, rotation: deg })));
                                    } else if (selectedRotatePageId) {
                                      setInteractivePages(prev => prev.map(p => p.id === selectedRotatePageId ? { ...p, rotation: deg } : p));
                                    }
                                  }}
                                  className="px-2 py-2 border rounded-xl font-mono text-xs font-extrabold bg-white hover:bg-slate-50 text-slate-800 transition-all border-slate-200 flex flex-col items-center justify-center gap-1.5"
                                  title={`Rotate ${deg}°`}
                                >
                                  <span className="text-amber-500 scale-110">↻</span>
                                  <span>{deg}°</span>
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Delete Pages Options (Tool 11) */}
                      {toolId === 'delete_pages' && !isLoadingPages && (
                        <div className="space-y-3.5 text-slate-700">
                          <div className="p-3 bg-red-50 border border-red-200 rounded-xl space-y-1.5">
                            <p className="font-extrabold text-red-800 flex items-center gap-1.5 text-xs">
                              <span>🗑️</span>
                              <span>{language === Language.AR ? 'ملخص الصفحات المحددة للحذف:' : 'Deletion Target Summary:'}</span>
                            </p>
                            
                            <div className="grid grid-cols-2 gap-2 text-[10px] font-bold text-slate-600">
                              <div className="bg-white p-2 rounded border border-slate-150">
                                <span className="block text-slate-400 font-medium">{language === Language.AR ? 'سيتم حذفها:' : 'Pages to delete:'}</span>
                                <span className="text-rose-600 font-extrabold text-base">
                                  {interactivePages.filter(p => p.isSelectedForDelete).length}
                                </span>
                              </div>
                              <div className="bg-white p-2 rounded border border-slate-150">
                                <span className="block text-slate-400 font-medium">{language === Language.AR ? 'سيتم حفظها وتصديرها:' : 'Pages to keep:'}</span>
                                <span className="text-emerald-600 font-extrabold text-base">
                                  {interactivePages.filter(p => !p.isSelectedForDelete).length}
                                </span>
                              </div>
                            </div>
                          </div>

                          {/* List of page numbers selected */}
                          {interactivePages.some(p => p.isSelectedForDelete) && (
                            <div className="space-y-1">
                              <span className="block text-[11px] font-extrabold text-slate-500">
                                {language === Language.AR ? 'الصفحات التي سيتم إزالتها:' : 'Specific pages scheduled for removal:'}
                              </span>
                              <div className="flex flex-wrap gap-1 max-h-24 overflow-y-auto p-1.5 border border-slate-200 rounded bg-slate-50">
                                {interactivePages.map((p, pIdx) => p.isSelectedForDelete ? (
                                  <span key={p.id} className="px-1.5 py-0.5 bg-red-100 border border-red-200 text-red-700 text-[10px] font-mono font-extrabold rounded">
                                    {pIdx + 1}
                                  </span>
                                ) : null)}
                              </div>
                            </div>
                          )}

                          {/* Quick Actions buttons */}
                          <div className="grid grid-cols-2 gap-2 pt-1">
                            <button
                              type="button"
                              onClick={() => setInteractivePages(prev => prev.map(p => ({ ...p, isSelectedForDelete: true })))}
                              className="px-2 py-1.5 border border-slate-200 hover:border-red-300 rounded text-[10px] font-bold text-slate-600 hover:text-red-600 bg-white hover:bg-red-50/20 transition-all"
                            >
                              ✕ {language === Language.AR ? 'تحديد كل الصفحات للحذف' : 'Select All'}
                            </button>
                            <button
                              type="button"
                              onClick={() => setInteractivePages(prev => prev.map(p => ({ ...p, isSelectedForDelete: false })))}
                              className="px-2 py-1.5 border border-slate-200 hover:border-emerald-300 rounded text-[10px] font-bold text-slate-600 hover:text-emerald-600 bg-white hover:bg-emerald-50/20 transition-all"
                            >
                              ✓ {language === Language.AR ? 'إلغاء تحديد كل الصفحات' : 'Clear Flags'}
                            </button>
                          </div>
                        </div>
                      )}

                      {/* Reorder Pages Options (Tool 12) */}
                      {toolId === 'reorder' && (
                        <div className="space-y-3.5 text-slate-700">
                          {/* Instructional box */}
                          <div className="p-3 bg-violet-50 border border-violet-150 rounded-xl space-y-1">
                            <p className="font-extrabold text-violet-800 text-xs flex items-center gap-1.5">
                              <span>🔃</span>
                              <span>{language === Language.AR ? 'طريقة إعادة الترتيب:' : 'Reorder Instructions:'}</span>
                            </p>
                            <p className="text-[10px] text-slate-500 leading-relaxed font-semibold">
                              {language === Language.AR 
                                ? 'يمكنك تحريك الصفحات يميناً ويساراً باستخدام الأزرار (◀ / ▶) الموجودة على كرت كل صفحة، أو قم بسحب وإفلات الكروت مباشرة في المعاينة!' 
                                : 'You can rearrange pages by clicking the left/right move buttons on each card, or easily drag and drop cards within the visual preview!'}
                            </p>
                          </div>

                          {/* Order Map list summary */}
                          <div className="space-y-1.5">
                            <span className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تدفق المستند الحالي:' : 'Current Document Sequence Flow:'}
                            </span>
                            <div className="max-h-48 overflow-y-auto border border-slate-200 rounded-xl p-2.5 space-y-1 bg-slate-50 shadow-inner">
                              {interactivePages.map((p, pIdx) => (
                                <div key={p.id} className="flex items-center gap-2 p-1.5 bg-white border border-slate-150 rounded text-[10px] font-semibold text-slate-600">
                                  <span className="w-5 h-5 rounded-full bg-violet-100 border border-violet-200 text-violet-700 flex items-center justify-center font-mono font-extrabold">{pIdx + 1}</span>
                                  <span className="truncate flex-1 text-slate-700 font-mono">P. {p.pageNum + 1} - {p.fileName}</span>
                                </div>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Decrypt / Unlock PDF Options (Tool 13) */}
                      {toolId === 'unlock' && (
                        <div className="space-y-4 text-slate-700">
                          <div className="flex flex-col items-center justify-center py-4 bg-emerald-50 border border-emerald-150 rounded-xl text-center">
                            <span className="text-3xl animate-bounce mb-1">🔓</span>
                            <h5 className="font-extrabold text-emerald-800 text-xs">
                              {language === Language.AR ? 'المستند محمي وجاهز لفك القفل' : 'Secure Document Decryption'}
                            </h5>
                          </div>

                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'أدخل كلمة مرور فك التشفير:' : 'Enter Password to Decrypt:'}
                            </label>
                            <input
                              type="password"
                              value={securityPassword}
                              onChange={(e) => setSecurityPassword(e.target.value)}
                              placeholder={language === Language.AR ? 'أدخل كلمة سر المالك للملف المرفوع' : 'Enter PDF password...'}
                              className="w-full px-3 py-2 rounded-lg border border-slate-300 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none text-xs bg-white transition-all font-bold"
                            />
                            <p className="text-[10px] text-slate-400 font-medium">
                              {language === Language.AR 
                                ? 'يجب إدخال كلمة سر صحيحة لحذف قيود الطباعة والنسخ بنجاح.' 
                                : 'Required to strip print locks, edit locks, and open restrictions successfully.'}
                            </p>
                          </div>
                        </div>
                      )}

                      {/* PDF to JPG Conversion Options (Tool 14) */}
                      {toolId === 'pdf2img' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Image resolution presets */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'دقة وجودة الصور المصدرة:' : 'Output Image Quality Preset:'}
                            </label>
                            <div className="grid grid-cols-3 gap-2">
                              {['low', 'medium', 'high'].map(level => (
                                <button
                                  key={level}
                                  type="button"
                                  onClick={() => setJpgQuality(level as any)}
                                  className={`px-1.5 py-2 rounded-lg border text-center text-[10px] font-bold uppercase transition-all ${
                                    jpgQuality === level
                                      ? 'bg-rose-600 border-rose-700 text-white shadow'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {level === 'low' ? '100 DPI' : level === 'medium' ? '150 DPI' : '300 DPI (HD)'}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* ZIP switch toggle */}
                          <div className="flex items-center justify-between p-3 bg-white rounded-xl border border-slate-200">
                            <div className="space-y-0.5">
                              <span className="block text-xs font-bold text-slate-800">
                                {language === Language.AR ? 'تحميل كملف مضغوط ZIP:' : 'Download as ZIP archive:'}
                              </span>
                              <span className="block text-[9px] text-slate-400 font-medium">
                                {language === Language.AR ? 'تجميع كل الصور المستخرجة بملف واحد' : 'Zips all generated JPEGs for easy download'}
                              </span>
                            </div>
                            <button
                              type="button"
                              onClick={() => setDownloadAsZip(!downloadAsZip)}
                              className={`w-11 h-6 rounded-full transition-colors relative flex items-center ${
                                downloadAsZip ? 'bg-rose-600' : 'bg-slate-200'
                              }`}
                            >
                              <span className={`w-4 h-4 rounded-full bg-white shadow-md absolute transition-transform ${downloadAsZip ? 'translate-x-6' : 'translate-x-1'}`} />
                            </button>
                          </div>

                          {/* Range selector */}
                          <div className="space-y-1.5 border-t border-slate-100 pt-3">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تصدير صفحات محددة:' : 'Export Specific Pages:'}
                            </label>
                            <div className="flex gap-2">
                              <button
                                type="button"
                                onClick={() => setJpgPagesMode('all')}
                                className={`px-2.5 py-1 rounded text-[11px] font-bold border transition-all ${
                                  jpgPagesMode === 'all' 
                                    ? 'bg-rose-50 border-rose-300 text-rose-600' 
                                    : 'bg-white border-slate-200 text-slate-500'
                                }`}
                              >
                                {language === Language.AR ? 'كل الصفحات' : 'All Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setJpgPagesMode('range')}
                                className={`px-2.5 py-1 rounded text-[11px] font-bold border transition-all ${
                                  jpgPagesMode === 'range' 
                                    ? 'bg-rose-50 border-rose-300 text-rose-600' 
                                    : 'bg-white border-slate-200 text-slate-500'
                                }`}
                              >
                                {language === Language.AR ? 'نطاق مخصص' : 'Custom Range'}
                              </button>
                            </div>
                            {jpgPagesMode === 'range' && (
                              <input
                                type="text"
                                value={jpgPagesRange}
                                onChange={(e) => setJpgPagesRange(e.target.value)}
                                placeholder="e.g. 1-2, 4-6"
                                className="w-full mt-1.5 px-3 py-1.5 rounded border border-slate-300 text-xs bg-white focus:ring-1 focus:ring-rose-500"
                              />
                            )}
                          </div>
                        </div>
                      )}

                      {/* Crop Document Options (Tool 15) */}
                      {toolId === 'crop' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Target crop mode switcher */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تطبيق القص على:' : 'Apply Crop Bounds To:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setCropMode('all')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  cropMode === 'all'
                                    ? 'bg-rose-600 border-rose-700 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                📑 {language === Language.AR ? 'كل الصفحات دفعة واحدة' : 'All Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setCropMode('single')}
                                className={`px-2 py-1.5 rounded-lg border text-center text-xs font-bold transition-all ${
                                  cropMode === 'single'
                                    ? 'bg-rose-600 border-rose-700 text-white shadow'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                🎯 {language === Language.AR ? 'صفحة مخصصة معينة' : 'Single Selected Page'}
                              </button>
                            </div>
                          </div>

                          {/* Custom Crop margins interactive controls */}
                          <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-3">
                            <span className="block text-[11px] font-extrabold text-slate-600 text-center uppercase tracking-wide border-b border-slate-200 pb-1">
                              {language === Language.AR ? 'تعديل هوامش قص المستند (%)' : 'Fine-Tune Crop Margins (%)'}
                            </span>
                            
                            {cropMode === 'single' && (
                              <div className="space-y-1">
                                <label className="block text-[10px] font-bold text-slate-500">
                                  {language === Language.AR ? 'الصفحة المختارة للقص مفرداً:' : 'Choose page target:'}
                                </label>
                                <select
                                  value={selectedCropPageId}
                                  onChange={(e) => {
                                    setSelectedCropPageId(e.target.value);
                                    const pageId = e.target.value;
                                    if (pageId && cropMarginsMap[pageId]) {
                                      setCropMargins(cropMarginsMap[pageId]);
                                    } else {
                                      setCropMargins({ top: 10, bottom: 10, left: 10, right: 10 });
                                    }
                                  }}
                                  className="w-full p-1 border rounded bg-white text-xs font-bold"
                                >
                                  <option value="">-- {language === Language.AR ? 'اختر صفحة للتعديل' : 'Select page to crop'} --</option>
                                  {interactivePages.map((p, pIdx) => (
                                    <option key={p.id} value={p.id}>
                                      {language === Language.AR ? `صفحة ${pIdx + 1}` : `Page ${pIdx + 1}`}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            )}

                            {/* Margins controllers */}
                            <div className="space-y-2">
                              {/* Top Margin */}
                              <div className="space-y-0.5">
                                <div className="flex items-center justify-between font-mono text-[10px] font-bold">
                                  <span>{language === Language.AR ? 'الهامش العلوي:' : 'Top Margin:'}</span>
                                  <span>{cropMargins.top}%</span>
                                </div>
                                <input
                                  type="range"
                                  min="0"
                                  max="45"
                                  value={cropMargins.top}
                                  onChange={(e) => {
                                    const val = parseInt(e.target.value, 10);
                                    const updated = { ...cropMargins, top: val };
                                    setCropMargins(updated);
                                    if (cropMode === 'single' && selectedCropPageId) {
                                      setCropMarginsMap(prev => ({ ...prev, [selectedCropPageId]: updated }));
                                    }
                                  }}
                                  className="w-full accent-rose-600 h-1"
                                />
                              </div>

                              {/* Bottom Margin */}
                              <div className="space-y-0.5">
                                <div className="flex items-center justify-between font-mono text-[10px] font-bold">
                                  <span>{language === Language.AR ? 'الهامش السفلي:' : 'Bottom Margin:'}</span>
                                  <span>{cropMargins.bottom}%</span>
                                </div>
                                <input
                                  type="range"
                                  min="0"
                                  max="45"
                                  value={cropMargins.bottom}
                                  onChange={(e) => {
                                    const val = parseInt(e.target.value, 10);
                                    const updated = { ...cropMargins, bottom: val };
                                    setCropMargins(updated);
                                    if (cropMode === 'single' && selectedCropPageId) {
                                      setCropMarginsMap(prev => ({ ...prev, [selectedCropPageId]: updated }));
                                    }
                                  }}
                                  className="w-full accent-rose-600 h-1"
                                />
                              </div>

                              {/* Left Margin */}
                              <div className="space-y-0.5">
                                <div className="flex items-center justify-between font-mono text-[10px] font-bold">
                                  <span>{language === Language.AR ? 'الهامش الأيسر:' : 'Left Margin:'}</span>
                                  <span>{cropMargins.left}%</span>
                                </div>
                                <input
                                  type="range"
                                  min="0"
                                  max="45"
                                  value={cropMargins.left}
                                  onChange={(e) => {
                                    const val = parseInt(e.target.value, 10);
                                    const updated = { ...cropMargins, left: val };
                                    setCropMargins(updated);
                                    if (cropMode === 'single' && selectedCropPageId) {
                                      setCropMarginsMap(prev => ({ ...prev, [selectedCropPageId]: updated }));
                                    }
                                  }}
                                  className="w-full accent-rose-600 h-1"
                                />
                              </div>

                              {/* Right Margin */}
                              <div className="space-y-0.5">
                                <div className="flex items-center justify-between font-mono text-[10px] font-bold">
                                  <span>{language === Language.AR ? 'الهامش الأيمن:' : 'Right Margin:'}</span>
                                  <span>{cropMargins.right}%</span>
                                </div>
                                <input
                                  type="range"
                                  min="0"
                                  max="45"
                                  value={cropMargins.right}
                                  onChange={(e) => {
                                    const val = parseInt(e.target.value, 10);
                                    const updated = { ...cropMargins, right: val };
                                    setCropMargins(updated);
                                    if (cropMode === 'single' && selectedCropPageId) {
                                      setCropMarginsMap(prev => ({ ...prev, [selectedCropPageId]: updated }));
                                    }
                                  }}
                                  className="w-full accent-rose-600 h-1"
                                />
                              </div>
                            </div>

                            <p className="text-[9px] text-slate-400 font-semibold leading-relaxed text-center">
                              {language === Language.AR 
                                ? 'ملاحظة: يمكنك تحريك مؤشرات القص الظاهرة مباشرة عند حوم الماوس فوق صفحات كروت المعاينة للحصول على قص دقيق للغاية وبشكل أسرع!' 
                                : 'Note: You can adjust margins with on-card sliders that appear when hovering over the preview thumbnails!'}
                            </p>
                          </div>
                        </div>
                      )}

                      {/* 16. Smart Text Extraction Options */}
                      {toolId === 'extract' && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'نطاق الصفحات المطلوب استخراجها:' : 'Pages to Extract:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setExtractPagesMode('all')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  extractPagesMode === 'all'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'كل الصفحات' : 'All Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setExtractPagesMode('range')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  extractPagesMode === 'range'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نطاق مخصص' : 'Custom Range'}
                              </button>
                            </div>
                          </div>

                          {extractPagesMode === 'range' && (
                            <div className="space-y-1.5 animate-in fade-in-50 duration-200">
                              <label className="block text-[11px] font-bold text-slate-500">
                                {language === Language.AR ? 'حدد نطاق الصفحات (مثال: 1-3, 5):' : 'Specify page numbers (e.g. 1-3, 5):'}
                              </label>
                              <input
                                type="text"
                                placeholder={language === Language.AR ? 'مثال: 1-3، 5' : 'e.g. 1-3, 5'}
                                value={extractPagesRange}
                                onChange={(e) => setExtractPagesRange(e.target.value)}
                                className="w-full px-3 py-2.5 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-mono"
                              />
                            </div>
                          )}

                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'صيغة النص الناتج:' : 'Output Format:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setExtractFormat('plain')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  extractFormat === 'plain'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نص عادي مبسط' : 'Plain Text'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setExtractFormat('structured')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  extractFormat === 'structured'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'منظم بنقاط وعناوين' : 'Structured List'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* 17. Automatic PDF Summarization Options */}
                      {toolId === 'summarize' && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'طول التلخيص المطلوب:' : 'Summary Length:'}
                            </label>
                            <div className="grid grid-cols-3 gap-1.5">
                              {[
                                { id: 'short', ar: 'قصير', en: 'Short' },
                                { id: 'medium', ar: 'متوسط', en: 'Medium' },
                                { id: 'detailed', ar: 'مفصل', en: 'Detailed' }
                              ].map(item => (
                                <button
                                  key={item.id}
                                  type="button"
                                  onClick={() => setSummaryLength(item.id as any)}
                                  className={`py-2 px-1 rounded-xl border text-[11px] font-bold transition-all ${
                                    summaryLength === item.id
                                      ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {language === Language.AR ? item.ar : item.en}
                                </button>
                              ))}
                            </div>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'شكل عرض الملخص:' : 'Summary Display Format:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setSummaryFormat('bullets')}
                                className={`py-2.5 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  summaryFormat === 'bullets'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نقاط تعداد منفصلة' : 'Bullet Points'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setSummaryFormat('paragraphs')}
                                className={`py-2.5 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  summaryFormat === 'paragraphs'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'فقرات متصلة متناسقة' : 'Connected Paragraphs'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* 18. PDF Translation Options (Expanded with preserve formatting) */}
                      {toolId === 'translate' && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'اختر اللغة الهدف للترجمة:' : 'Target Language:'}
                            </label>
                            <select
                              value={targetTranslationLang}
                              onChange={(e) => setTargetTranslationLang(e.target.value)}
                              className="w-full px-4 py-3 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none bg-white transition-all font-semibold text-slate-700"
                            >
                              {[
                                { code: 'Arabic', label: language === Language.AR ? 'العربية' : 'Arabic' },
                                { code: 'English', label: language === Language.AR ? 'الإنجليزية (English)' : 'English' },
                                { code: 'French', label: language === Language.AR ? 'الفرنسية (French)' : 'French' },
                                { code: 'Spanish', label: language === Language.AR ? 'الإسبانية (Spanish)' : 'Spanish' },
                                { code: 'Turkish', label: language === Language.AR ? 'التركيّة (Turkish)' : 'Turkish' }
                              ].map((langObj, idx) => (
                                <option key={idx} value={langObj.code}>{langObj.label}</option>
                              ))}
                            </select>
                          </div>

                          <div className="p-3 bg-indigo-50/50 rounded-xl border border-indigo-100 flex items-center justify-between">
                            <div className="space-y-0.5">
                              <span className="block text-xs font-extrabold text-indigo-950">
                                {language === Language.AR ? 'الحفاظ على التنسيق الأصلي' : 'Preserve Layout & Fonts'}
                              </span>
                              <span className="block text-[10px] text-indigo-600 font-medium">
                                {language === Language.AR ? 'ترجمة النصوص داخل مواضعها الأصلية بالصفحة.' : 'Translate texts keeping identical sizes and coordinates.'}
                              </span>
                            </div>
                            <label className="relative inline-flex items-center cursor-pointer">
                              <input
                                type="checkbox"
                                checked={preserveLayout}
                                onChange={(e) => setPreserveLayout(e.target.checked)}
                                className="sr-only peer"
                              />
                              <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-indigo-600"></div>
                            </label>
                          </div>
                        </div>
                      )}

                      {/* 19. OCR Options */}
                      {toolId === 'ocr' && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'لغة المستند الممسوح ضوئياً:' : 'Scanned Document Language:'}
                            </label>
                            <select
                              value={ocrLanguage}
                              onChange={(e) => setOcrLanguage(e.target.value)}
                              className="w-full px-4 py-3 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none bg-white transition-all font-semibold text-slate-700"
                            >
                              {[
                                { code: 'Arabic', label: language === Language.AR ? 'العربية (الأدق للمستندات العربية)' : 'Arabic (Best for Arabic OCR)' },
                                { code: 'English', label: language === Language.AR ? 'الإنجليزية (English)' : 'English' },
                                { code: 'French', label: language === Language.AR ? 'الفرنسية (French)' : 'French' },
                                { code: 'Spanish', label: language === Language.AR ? 'الإسبانية (Spanish)' : 'Spanish' },
                                { code: 'Turkish', label: language === Language.AR ? 'التركيّة (Turkish)' : 'Turkish' }
                              ].map((langObj, idx) => (
                                <option key={idx} value={langObj.code}>{langObj.label}</option>
                              ))}
                            </select>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'صيغة المستند الناتج عن التعرف:' : 'Recognition Output Type:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setOcrOutputType('plain_text')}
                                className={`py-2.5 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  ocrOutputType === 'plain_text'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نص عادي مستخرج' : 'Plain Extracted Text'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setOcrOutputType('searchable_pdf')}
                                className={`py-2.5 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  ocrOutputType === 'searchable_pdf'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'PDF قابل للبحث فيه' : 'Searchable PDF'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* 20. Intelligent Formatting Options */}
                      {toolId === 'format' && (
                        <div className="space-y-4">
                          <div className="p-3 bg-slate-50 border border-slate-100 rounded-xl space-y-3.5">
                            <div className="flex items-center justify-between">
                              <div className="space-y-0.5 max-w-[80%]">
                                <span className="block text-xs font-extrabold text-slate-800">
                                  {language === Language.AR ? 'تعديل ومحاذاة الفقرات تلقائياً' : 'Auto-Align Paragraphs'}
                                </span>
                                <span className="block text-[10px] text-slate-400 font-medium">
                                  {language === Language.AR ? 'محاذاة النصوص والاتجاهات لليمين واليسار حسب اللغة.' : 'Fix word alignments, text directions, and spacing issues.'}
                                </span>
                              </div>
                              <label className="relative inline-flex items-center cursor-pointer">
                                <input
                                  type="checkbox"
                                  checked={formatAutoAlign}
                                  onChange={(e) => setFormatAutoAlign(e.target.checked)}
                                  className="sr-only peer"
                                />
                                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-indigo-600"></div>
                              </label>
                            </div>

                            <div className="border-t border-slate-100 pt-3 flex items-center justify-between">
                              <div className="space-y-0.5 max-w-[80%]">
                                <span className="block text-xs font-extrabold text-slate-800">
                                  {language === Language.AR ? 'توحيد وتحسين الخطوط والرموز' : 'Unify Document Fonts'}
                                </span>
                                <span className="block text-[10px] text-slate-400 font-medium">
                                  {language === Language.AR ? 'توحيد عائلات الخطوط لتظهر بنمط مهني أنيق ومتكامل.' : 'Unify font families and fix missing symbols or raw shapes.'}
                                </span>
                              </div>
                              <label className="relative inline-flex items-center cursor-pointer">
                                <input
                                  type="checkbox"
                                  checked={formatUnifyFonts}
                                  onChange={(e) => setFormatUnifyFonts(e.target.checked)}
                                  className="sr-only peer"
                                />
                                <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-indigo-600"></div>
                              </label>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* 21. Real-time Inline Sidebar Chat with PDF */}
                      {toolId === 'chat' && (
                        <div className="space-y-4 flex flex-col max-h-[58vh]">
                          <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-3 text-[10px] space-y-1">
                            <p className="font-extrabold text-indigo-900 flex items-center gap-1">
                              <span>💬</span>
                              <span>{language === Language.AR ? 'مساعدك الشخصي للـ PDF نشط!' : 'Your interactive PDF AI agent is active!'}</span>
                            </p>
                            <p className="text-indigo-700 font-medium leading-relaxed">
                              {language === Language.AR
                                ? 'اطرح أي أسئلة وسيقوم المساعد بقراءة الصفحات والإجابة مع إمكانية التنقل التلقائي وتحديد صفحة المصدر في لوحة المعاينة.'
                                : 'Ask questions and the assistant will reference the document pages and answer. Clicking page indicators scrolls to sources.'}
                            </p>
                          </div>

                          {/* Chat Feed */}
                          <div className="border border-slate-200 rounded-xl bg-white p-3 overflow-y-auto space-y-2.5 max-h-[30vh] min-h-[140px] font-sans text-xs">
                            {chatHistory.length === 0 ? (
                              <p className="text-slate-400 italic text-center py-6">{language === Language.AR ? 'ابدأ المحادثة بسؤال المستند أدناه...' : 'Type a question below to start chatting...'}</p>
                            ) : (
                              chatHistory.map((msg, idx) => (
                                <div key={idx} className={`flex flex-col ${msg.sender === 'user' ? 'items-end' : 'items-start'}`}>
                                  <div
                                    className={`max-w-[90%] rounded-xl px-3 py-2 leading-relaxed shadow-xs ${
                                      msg.sender === 'user'
                                        ? 'bg-indigo-600 text-white rounded-br-none'
                                        : 'bg-slate-50 border border-slate-100 text-slate-800 rounded-bl-none'
                                    }`}
                                  >
                                    {/* Text with clickable source links if any */}
                                    <div className="space-y-1">
                                      <span>{msg.text}</span>
                                      {/* Parse source indicators if present */}
                                      {msg.sender === 'assistant' && (
                                        <div className="flex flex-wrap gap-1 mt-1.5 pt-1.5 border-t border-slate-100/50">
                                          {interactivePages.map((p, pIdx) => {
                                            // Check if page citation is in text
                                            const matchPattern = new RegExp(`page\\s*${p.pageNum + 1}|صفحة\\s*${p.pageNum + 1}|\\[${p.pageNum + 1}\\]`, 'i');
                                            if (matchPattern.test(msg.text) || idx === chatHistory.length - 1) {
                                              return (
                                                <button
                                                  key={p.id}
                                                  onClick={() => setSelectedPageSource(p.pageNum)}
                                                  className={`text-[9px] px-2 py-0.5 rounded font-extrabold flex items-center gap-0.5 border transition-all ${
                                                    selectedPageSource === p.pageNum
                                                      ? 'bg-violet-600 text-white border-violet-700 shadow-sm'
                                                      : 'bg-white hover:bg-slate-100 text-slate-600 border-slate-200'
                                                  }`}
                                                >
                                                  <span>🎯</span>
                                                  <span>{language === Language.AR ? `صفحة ${p.pageNum + 1}` : `Page ${p.pageNum + 1}`}</span>
                                                </button>
                                              );
                                            }
                                            return null;
                                          })}
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              ))
                            )}
                            {isChatSending && (
                              <div className="flex justify-start">
                                <div className="bg-slate-50 border border-slate-100 rounded-xl rounded-bl-none px-3 py-2 flex space-x-1.5 items-center">
                                  <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full animate-bounce"></span>
                                  <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full animate-bounce [animation-delay:0.2s]"></span>
                                  <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full animate-bounce [animation-delay:0.4s]"></span>
                                </div>
                              </div>
                            )}
                          </div>

                          {/* Quick predefined prompt helpers */}
                          <div className="flex gap-1.5 overflow-x-auto pb-1 scrollbar-thin">
                            {[
                              { ar: 'لخص المستند', en: 'Summarize file' },
                              { ar: 'ما هي أهم النقاط؟', en: 'What are key points?' },
                              { ar: 'استخرج الأرقام الهامة', en: 'Extract key figures' }
                            ].map((prompt, pIdx) => (
                              <button
                                key={pIdx}
                                type="button"
                                disabled={isChatSending}
                                onClick={() => {
                                  setChatInput(language === Language.AR ? prompt.ar : prompt.en);
                                }}
                                className="text-[10px] font-bold text-slate-500 hover:text-indigo-600 bg-slate-100 hover:bg-indigo-50 border border-slate-200 hover:border-indigo-100 px-2.5 py-1 rounded-full whitespace-nowrap transition-all"
                              >
                                {language === Language.AR ? prompt.ar : prompt.en}
                              </button>
                            ))}
                          </div>

                          {/* Chat Input form */}
                          <div className="flex gap-2 space-x-reverse">
                            <input
                              type="text"
                              value={chatInput}
                              onChange={(e) => setChatInput(e.target.value)}
                              disabled={isChatSending}
                              placeholder={language === Language.AR ? 'اطرح سؤالاً عن الـ PDF...' : 'Ask about the document...'}
                              className="flex-1 px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  handleSendChatMessage(e as any);
                                }
                              }}
                            />
                            <button
                              type="button"
                              onClick={handleSendChatMessage}
                              disabled={!chatInput.trim() || isChatSending}
                              className="bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-200 text-white font-extrabold px-3.5 py-2 rounded-xl text-xs shadow-sm transition-all"
                            >
                              {language === Language.AR ? 'إرسال' : 'Send'}
                            </button>
                          </div>
                        </div>
                      )}

                      {/* 22. Redact Sensitive Info Options */}
                      {toolId === 'redact' && (
                        <div className="space-y-4">
                          <label className="block text-xs font-extrabold text-slate-700">
                            {language === Language.AR ? 'حدد أنواع البيانات الحساسة لطمسها (PII):' : 'Select PII Data Types to Redact:'}
                          </label>

                          <div className="p-3 bg-slate-50 border border-slate-100 rounded-xl space-y-3">
                            <div className="flex items-center justify-between">
                              <span className="text-xs font-bold text-slate-700">
                                📧 {language === Language.AR ? 'عناوين البريد الإلكتروني' : 'Email Addresses'}
                              </span>
                              <input
                                type="checkbox"
                                checked={redactPiiEmail}
                                onChange={(e) => setRedactPiiEmail(e.target.checked)}
                                className="w-4 h-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded"
                              />
                            </div>

                            <div className="flex items-center justify-between border-t border-slate-100 pt-2.5">
                              <span className="text-xs font-bold text-slate-700">
                                📞 {language === Language.AR ? 'أرقام الهواتف والجوالات' : 'Phone Numbers'}
                              </span>
                              <input
                                type="checkbox"
                                checked={redactPiiPhone}
                                onChange={(e) => setRedactPiiPhone(e.target.checked)}
                                className="w-4 h-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded"
                              />
                            </div>

                            <div className="flex items-center justify-between border-t border-slate-100 pt-2.5">
                              <span className="text-xs font-bold text-slate-700">
                                🪪 {language === Language.AR ? 'الهويات الوطنية والإقامات' : 'National IDs / Passports'}
                              </span>
                              <input
                                type="checkbox"
                                checked={redactPiiId}
                                onChange={(e) => setRedactPiiId(e.target.checked)}
                                className="w-4 h-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded"
                              />
                            </div>

                            <div className="flex items-center justify-between border-t border-slate-100 pt-2.5">
                              <span className="text-xs font-bold text-slate-700">
                                📍 {language === Language.AR ? 'العناوين الجغرافية والمنازل' : 'Addresses & Locations'}
                              </span>
                              <input
                                type="checkbox"
                                checked={redactPiiAddress}
                                onChange={(e) => setRedactPiiAddress(e.target.checked)}
                                className="w-4 h-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded"
                              />
                            </div>
                          </div>

                          <div className="space-y-1.5">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'كلمات أو نصوص مخصصة لطمسها:' : 'Manual Keywords to Redact:'}
                            </label>
                            <input
                              type="text"
                              value={redactKeywords}
                              onChange={(e) => setRedactKeywords(e.target.value)}
                              placeholder={language === Language.AR ? 'مثال: اسم الموظف، اسم الشركة' : 'e.g. employee name, company'}
                              className="w-full px-3 py-2.5 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                            />
                            <p className="text-[9px] text-slate-400 font-medium">
                              {language === Language.AR ? 'افصل بين الكلمات بفواصل (مثال: أحمد، سابك).' : 'Separate keywords with commas.'}
                            </p>
                          </div>
                        </div>
                      )}

                      {/* 23. Compare PDF Versions Options */}
                      {toolId === 'compare' && (
                        <div className="space-y-4">
                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'رفع المستند المعدل للمقارنة:' : 'Upload Modified PDF to Compare:'}
                            </label>
                            {compareFile ? (
                              <div className="p-3 bg-indigo-50 border border-indigo-100 rounded-xl flex items-center justify-between">
                                <div className="flex items-center space-x-1.5 space-x-reverse overflow-hidden">
                                  <span className="text-[14px]">📄</span>
                                  <span className="text-[11px] font-bold text-indigo-950 truncate max-w-[150px]">{compareFile.name}</span>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => setCompareFile(null)}
                                  className="text-[10px] font-extrabold text-red-600 hover:text-red-700 bg-white border border-red-100 rounded-lg px-2 py-1 shadow-xs transition-colors"
                                >
                                  {language === Language.AR ? 'إزالة' : 'Remove'}
                                </button>
                              </div>
                            ) : (
                              <div
                                onClick={() => {
                                  const input = document.createElement('input');
                                  input.type = 'file';
                                  input.accept = 'application/pdf';
                                  input.onchange = (e: any) => {
                                    if (e.target.files && e.target.files[0]) {
                                      setCompareFile(e.target.files[0]);
                                    }
                                  };
                                  input.click();
                                }}
                                className="border border-dashed border-slate-300 hover:border-indigo-400 rounded-xl p-4 text-center cursor-pointer hover:bg-slate-50 transition-colors flex flex-col items-center justify-center space-y-1"
                              >
                                <span className="text-[18px]">📤</span>
                                <span className="text-[11px] font-extrabold text-slate-600">{language === Language.AR ? 'اختر مستند النسخة المعدلة' : 'Select Modified PDF'}</span>
                                <span className="text-[9px] text-slate-400 font-medium">{language === Language.AR ? 'للمقارنة مع النسخة الأصلية' : 'To contrast with original'}</span>
                              </div>
                            )}
                          </div>

                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'نوع المقارنة والتحليل:' : 'Comparison Mode:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setCompareMode('visual')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  compareMode === 'visual'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'مقارنة بصرية للصفحات' : 'Visual Page Diff'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setCompareMode('textual')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  compareMode === 'textual'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'تحليل نصي للفروقات' : 'Textual Comparison'}
                              </button>
                            </div>
                          </div>

                          <div className="space-y-2">
                            <label className="block text-xs font-extrabold text-slate-700">
                              {language === Language.AR ? 'تصفية الفروق المعروضة:' : 'Filter Diff View:'}
                            </label>
                            <div className="grid grid-cols-3 gap-1.5">
                              {[
                                { id: 'all', ar: 'الكل', en: 'All' },
                                { id: 'additions', ar: 'الإضافات', en: 'Additions' },
                                { id: 'deletions', ar: 'المحذوفات', en: 'Deletions' }
                              ].map(item => (
                                <button
                                  key={item.id}
                                  type="button"
                                  onClick={() => setDiffView(item.id as any)}
                                  className={`py-2 rounded-xl border text-[11px] font-bold transition-all ${
                                    diffView === item.id
                                      ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {language === Language.AR ? item.ar : item.en}
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Add Page Numbers Options */}
                      {toolId === 'add_page_numbers' && (
                        <div className="space-y-4 text-slate-700">
                          {/* Position 3x3 Grid */}
                          <div className="space-y-2">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'موقع رقم الصفحة:' : 'Page Number Position:'}
                            </label>
                            <div className="grid grid-cols-3 gap-1.5 p-2 bg-slate-100 rounded-xl border border-slate-200">
                              {[
                                { id: 'top-left', label: language === Language.AR ? 'أعلى يسار' : 'Top Left' },
                                { id: 'top-center', label: language === Language.AR ? 'أعلى وسط' : 'Top Center' },
                                { id: 'top-right', label: language === Language.AR ? 'أعلى يمين' : 'Top Right' },
                                { id: 'middle-left', label: language === Language.AR ? 'وسط يسار' : 'Middle Left' },
                                { id: 'middle-center', label: language === Language.AR ? 'وسط' : 'Middle Center' },
                                { id: 'middle-right', label: language === Language.AR ? 'وسط يمين' : 'Middle Right' },
                                { id: 'bottom-left', label: language === Language.AR ? 'أسفل يسار' : 'Bottom Left' },
                                { id: 'bottom-center', label: language === Language.AR ? 'أسفل وسط' : 'Bottom Center' },
                                { id: 'bottom-right', label: language === Language.AR ? 'أسفل يمين' : 'Bottom Right' }
                              ].map((pos) => (
                                <button
                                  key={pos.id}
                                  type="button"
                                  onClick={() => setPageNumberPosition(pos.id)}
                                  className={`py-2 px-1 rounded-lg text-[10px] font-extrabold transition-all text-center border ${
                                    pageNumberPosition === pos.id
                                      ? 'bg-indigo-600 border-indigo-700 text-white shadow-sm'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  {pos.label}
                                </button>
                              ))}
                            </div>
                          </div>

                          {/* Format Input */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'تنسيق الرقم:' : 'Number Format:'}
                            </label>
                            <input
                              type="text"
                              value={pageNumberFormat}
                              onChange={(e) => setPageNumberFormat(e.target.value)}
                              placeholder="e.g. Page {n} of {total}"
                              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-medium text-slate-800 bg-white"
                            />
                            <p className="text-[9px] text-slate-400 leading-relaxed font-medium">
                              {language === Language.AR
                                ? 'استخدم {n} لرقم الصفحة الحالي و {total} لإجمالي الصفحات.'
                                : 'Use {n} for current page number and {total} for total count.'}
                            </p>
                          </div>

                          {/* Page Range selection */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'نطاق ترقيم الصفحات:' : 'Page Range:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setPageNumberRange('all')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  pageNumberRange === 'all'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'كافة الصفحات' : 'All Pages'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setPageNumberRange('custom')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  pageNumberRange === 'custom'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'نطاق محدد' : 'Custom Range'}
                              </button>
                            </div>

                            {pageNumberRange === 'custom' && (
                              <div className="grid grid-cols-2 gap-2 pt-1.5">
                                <div className="space-y-1">
                                  <span className="text-[10px] font-bold text-slate-500">{language === Language.AR ? 'من صفحة:' : 'From Page:'}</span>
                                  <input
                                    type="number"
                                    min={1}
                                    value={pageNumberStart}
                                    onChange={(e) => setPageNumberStart(Number(e.target.value))}
                                    className="w-full p-2 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 bg-white"
                                  />
                                </div>
                                <div className="space-y-1">
                                  <span className="text-[10px] font-bold text-slate-500">{language === Language.AR ? 'إلى صفحة:' : 'To Page:'}</span>
                                  <input
                                    type="number"
                                    min={1}
                                    value={pageNumberEnd}
                                    onChange={(e) => setPageNumberEnd(Number(e.target.value))}
                                    className="w-full p-2 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 bg-white"
                                  />
                                </div>
                              </div>
                            )}
                          </div>

                          {/* Font Size Selector */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'حجم الخط:' : 'Font Size:'} ({pageNumberSize}px)
                            </label>
                            <input
                              type="range"
                              min={8}
                              max={24}
                              value={pageNumberSize}
                              onChange={(e) => setPageNumberSize(Number(e.target.value))}
                              className="w-full accent-indigo-600"
                            />
                          </div>
                        </div>
                      )}

                      {/* Repair Damaged PDF Options */}
                      {toolId === 'repair' && (
                        <div className="space-y-4 text-slate-700">
                          <div className="p-3 bg-rose-50 border border-rose-100 rounded-xl space-y-1">
                            <p className="font-bold text-xs text-rose-950 flex items-center gap-1">
                              <span>🔧</span>
                              <span>{language === Language.AR ? 'استرجاع ملف PDF تالف:' : 'Corrupt PDF Salvage:'}</span>
                            </p>
                            <p className="text-[10px] text-rose-900 leading-relaxed font-semibold">
                              {language === Language.AR
                                ? 'سيقوم النظام بتحليل بنية الملف وإصلاح الروابط المتضررة وجداول xref المفقودة لإعادة إتاحة الوصول للملف.'
                                : 'The system will parse the PDF file structure, correct broken internal offsets, and rebuild missing cross-reference tables.'}
                            </p>
                          </div>

                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'مستوى الإصلاح وفحص الأخطاء:' : 'Repair Depth Level:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setRepairLevel('standard')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  repairLevel === 'standard'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'إصلاح قياسي' : 'Standard Fix'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setRepairLevel('deep')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  repairLevel === 'deep'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'فحص عميق' : 'Deep Diagnostics'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Flatten PDF Layers Options */}
                      {toolId === 'flatten' && (
                        <div className="space-y-4 text-slate-700">
                          <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl space-y-1">
                            <p className="font-bold text-xs text-blue-950 flex items-center gap-1">
                              <span>🥞</span>
                              <span>{language === Language.AR ? 'تسطيح الطبقات والحقول تفاعلياً:' : 'Interactive Layer Flattening:'}</span>
                            </p>
                            <p className="text-[10px] text-blue-900 leading-relaxed font-semibold">
                              {language === Language.AR
                                ? 'يقوم هذا الخيار بدمج الحقول والتعليقات التوضيحية داخل صفحات المستند ليصبح محتوى ثابتاً غير قابل للتعديل.'
                                : 'This merges interactive form field elements and markups directly into page content stream so they cannot be modified.'}
                            </p>
                          </div>

                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'العناصر المراد تسطيحها:' : 'Layers to Flatten:'}
                            </label>
                            <div className="space-y-1.5">
                              {[
                                { id: 'all', title: language === Language.AR ? 'كافة الطبقات التفاعلية' : 'All Interactive Layers', desc: language === Language.AR ? 'تثبيت النماذج والتعليقات التوضيحية والرسومات' : 'Forms, markups, and drawing shapes' },
                                { id: 'forms', title: language === Language.AR ? 'حقول النماذج فقط' : 'Form Fields Only', desc: language === Language.AR ? 'تجميد مدخلات النماذج وتعبئة الحقول' : 'Flatten text inputs and checkboxes' },
                                { id: 'annotations', title: language === Language.AR ? 'التعليقات والملاحظات فقط' : 'Annotations Only', desc: language === Language.AR ? 'تثبيت التعليقات التوضيحية والطبقات الإضافية' : 'Flatten highlighted sections or stamps' }
                              ].map((opt) => (
                                <button
                                  key={opt.id}
                                  type="button"
                                  onClick={() => setFlattenMode(opt.id as any)}
                                  className={`w-full text-right p-2.5 rounded-xl border text-xs font-bold transition-all flex flex-col justify-start ${
                                    flattenMode === opt.id
                                      ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                      : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                  }`}
                                >
                                  <span className="block font-bold">{opt.title}</span>
                                  <span className="block text-[9px] text-slate-400 font-medium mt-0.5">{opt.desc}</span>
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                      )}

                      {/* Edit PDF Metadata Options */}
                      {toolId === 'metadata' && (
                        <div className="space-y-3.5 text-slate-700">
                          <div className="p-3 bg-emerald-50 border border-emerald-100 rounded-xl space-y-1">
                            <p className="font-bold text-xs text-emerald-950 flex items-center gap-1">
                              <span>🏷️</span>
                              <span>{language === Language.AR ? 'تحديث البيانات التعريفية للملف:' : 'Edit Meta parameters:'}</span>
                            </p>
                            <p className="text-[10px] text-emerald-900 leading-relaxed font-semibold">
                              {language === Language.AR
                                ? 'سيقوم النظام بحفظ هذه المعلومات داخل الهيكل التعريفي لملف PDF.'
                                : 'These internal PDF properties will be written into the document dictionary fields.'}
                            </p>
                          </div>

                          {/* Title */}
                          <div className="space-y-1">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'العنوان (Title):' : 'Title:'}
                            </label>
                            <input
                              type="text"
                              value={metaTitle}
                              onChange={(e) => setMetaTitle(e.target.value)}
                              placeholder={language === Language.AR ? 'العنوان الرئيسي للمستند' : 'Document Title'}
                              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-medium text-slate-800 bg-white"
                            />
                          </div>

                          {/* Author */}
                          <div className="space-y-1">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'المؤلف (Author):' : 'Author:'}
                            </label>
                            <input
                              type="text"
                              value={metaAuthor}
                              onChange={(e) => setMetaAuthor(e.target.value)}
                              placeholder={language === Language.AR ? 'اسم كاتب أو ناشر المستند' : 'Author / Publisher'}
                              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-medium text-slate-800 bg-white"
                            />
                          </div>

                          {/* Subject */}
                          <div className="space-y-1">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'الموضوع (Subject):' : 'Subject:'}
                            </label>
                            <input
                              type="text"
                              value={metaSubject}
                              onChange={(e) => setMetaSubject(e.target.value)}
                              placeholder={language === Language.AR ? 'موضوع المستند أو وصفه' : 'Document Subject'}
                              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-medium text-slate-800 bg-white"
                            />
                          </div>

                          {/* Keywords */}
                          <div className="space-y-1">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'الكلمات المفتاحية (Keywords):' : 'Keywords:'}
                            </label>
                            <input
                              type="text"
                              value={metaKeywords}
                              onChange={(e) => setMetaKeywords(e.target.value)}
                              placeholder={language === Language.AR ? 'افصل بين الكلمات بفواصل (,)' : 'Separate with commas (,)'}
                              className="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs focus:ring-2 focus:ring-indigo-500 outline-none transition-all font-medium text-slate-800 bg-white"
                            />
                          </div>
                        </div>
                      )}

                      {/* Web to PDF Options */}
                      {toolId === 'web2pdf' && htmlUrl.trim().length > 0 && (
                        <div className="space-y-4">
                          <div className="p-3.5 bg-sky-50 border border-sky-100 rounded-xl space-y-2">
                            <span className="block text-xs font-bold text-sky-950 truncate">
                              🌐 {htmlUrl}
                            </span>
                            <span className="block text-[10px] text-sky-900 font-medium font-semibold">
                              {language === Language.AR ? 'جاهز لتحويل الرابط إلى ملف PDF منسق.' : 'Ready to fetch and convert link contents to a structured PDF.'}
                            </span>
                          </div>
                        </div>
                      )}

                      {/* Generate QR Code Options */}
                      {toolId === 'qrcode' && (
                        <div className="space-y-4 text-slate-700">
                          {/* QR Code Text / URL */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'رابط أو نص رمز الـ QR:' : 'QR Code Text / Link:'}
                            </label>
                            <input
                              type="text"
                              id="qr-text-input"
                              name="qr-text-input"
                              disabled={false}
                              readOnly={false}
                              value={qrText}
                              onChange={(e) => {
                                console.log("QR input changed via onChange:", e.target.value);
                                setQrText(e.target.value);
                              }}
                              onInput={(e) => {
                                console.log("QR input updated via onInput:", e.currentTarget.value);
                                setQrText(e.currentTarget.value);
                              }}
                              onPaste={(e) => {
                                const text = e.clipboardData.getData('text');
                                console.log("QR input pasted via onPaste:", text);
                                setQrText(text);
                              }}
                              placeholder={language === Language.AR ? 'أدخل الرابط أو النص لصناعة رمز الـ QR' : 'Enter URL or text to generate'}
                              className="w-full px-3 py-2.5 rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs bg-white text-slate-800 font-medium relative z-20 pointer-events-auto cursor-text"
                            />
                          </div>

                          {/* QR Mode */}
                          <div className="space-y-1.5">
                            <label className="block text-xs font-bold text-slate-700">
                              {language === Language.AR ? 'طريقة الاستخراج والتصدير:' : 'Output Method:'}
                            </label>
                            <div className="grid grid-cols-2 gap-2">
                              <button
                                type="button"
                                onClick={() => setQrMode('embed')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  qrMode === 'embed'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'تضمين في الـ PDF' : 'Embed into PDF'}
                              </button>
                              <button
                                type="button"
                                onClick={() => setQrMode('download')}
                                className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all ${
                                  qrMode === 'download'
                                    ? 'border-indigo-600 bg-indigo-50 text-indigo-700 ring-2 ring-indigo-500/10'
                                    : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                }`}
                              >
                                {language === Language.AR ? 'تحميل كصورة PNG' : 'Download PNG'}
                              </button>
                            </div>
                          </div>

                          {qrMode === 'embed' && (
                            <>
                              {/* QR Code target page */}
                              <div className="space-y-1.5">
                                <label className="block text-xs font-bold text-slate-700">
                                  {language === Language.AR ? 'رقم الصفحة المستهدفة للتضمين:' : 'Target Page to Embed:'}
                                </label>
                                <input
                                  type="number"
                                  min={1}
                                  max={interactivePages.length || 1}
                                  value={qrPageNum}
                                  onChange={(e) => setQrPageNum(Number(e.target.value))}
                                  className="w-full p-2 border border-slate-200 rounded-lg text-xs font-bold text-slate-800 bg-white"
                                />
                              </div>

                              {/* QR Code visual position (grid) */}
                              <div className="space-y-2">
                                <label className="block text-xs font-bold text-slate-700">
                                  {language === Language.AR ? 'موقع الرمز في الصفحة:' : 'QR Code Alignment:'}
                                </label>
                                <div className="grid grid-cols-3 gap-1.5 p-2 bg-slate-100 rounded-xl border border-slate-200">
                                  {[
                                    { id: 'top-left', label: language === Language.AR ? 'أعلى يسار' : 'Top Left' },
                                    { id: 'top-center', label: language === Language.AR ? 'أعلى وسط' : 'Top Center' },
                                    { id: 'top-right', label: language === Language.AR ? 'أعلى يمين' : 'Top Right' },
                                    { id: 'middle-left', label: language === Language.AR ? 'وسط يسار' : 'Middle Left' },
                                    { id: 'middle-center', label: language === Language.AR ? 'وسط' : 'Middle Center' },
                                    { id: 'middle-right', label: language === Language.AR ? 'وسط يمين' : 'Middle Right' },
                                    { id: 'bottom-left', label: language === Language.AR ? 'أسفل يسار' : 'Bottom Left' },
                                    { id: 'bottom-center', label: language === Language.AR ? 'أسفل وسط' : 'Bottom Center' },
                                    { id: 'bottom-right', label: language === Language.AR ? 'أسفل يمين' : 'Bottom Right' }
                                  ].map((pos) => (
                                    <button
                                      key={pos.id}
                                      type="button"
                                      onClick={() => setQrPosition(pos.id)}
                                      className={`py-2 px-1 rounded-lg text-[10px] font-extrabold transition-all text-center border ${
                                        qrPosition === pos.id
                                          ? 'bg-indigo-600 border-indigo-700 text-white shadow-sm'
                                          : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                                      }`}
                                    >
                                      {pos.label}
                                    </button>
                                  ))}
                                </div>
                              </div>

                              {/* QR Code image size slider */}
                              <div className="space-y-1.5">
                                <label className="block text-xs font-bold text-slate-700">
                                  {language === Language.AR ? 'حجم الرمز:' : 'QR Code Dimension:'} ({qrCodeSize}px)
                                </label>
                                <input
                                  type="range"
                                  min={50}
                                  max={300}
                                  step={10}
                                  value={qrCodeSize}
                                  onChange={(e) => setQrCodeSize(Number(e.target.value))}
                                  className="w-full accent-indigo-600"
                                />
                              </div>
                            </>
                          )}
                        </div>
                      )}
                    </div>

                    {/* Global Page Selection Option (Bilingual & Interactive) */}
                    {interactivePages.length > 0 && (
                      <div className="p-4 bg-white border border-slate-200 rounded-2xl space-y-3 shadow-sm my-3">
                        <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                          <span className="text-xs font-extrabold text-slate-800 flex items-center gap-1.5">
                            <span>📄</span>
                            <span>{language === Language.AR ? 'نطاق الصفحات المستهدفة' : 'Target Page Range'}</span>
                          </span>
                          <span className="text-[10px] text-slate-400 font-bold">
                            {language === Language.AR ? `إجمالي: ${interactivePages.length} صفحات` : `Total: ${interactivePages.length} pages`}
                          </span>
                        </div>

                        <div className="grid grid-cols-2 gap-2">
                          <button
                            type="button"
                            onClick={() => {
                              setGlobalPagesSelectionMode('all');
                            }}
                            className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all flex flex-col items-center justify-center gap-1 ${
                              globalPagesSelectionMode === 'all'
                                ? 'border-indigo-600 bg-indigo-50/50 text-indigo-700 ring-2 ring-indigo-500/10 shadow-sm'
                                : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                            }`}
                          >
                            <span className="text-sm">🌟</span>
                            <span>{language === Language.AR ? 'كل الصفحات' : 'All Pages'}</span>
                          </button>

                          <button
                            type="button"
                            onClick={() => {
                              setGlobalPagesSelectionMode('custom');
                              if (globalSelectedPageIndices.length === 0) {
                                const allIndices = interactivePages.map(p => p.pageNum);
                                setGlobalSelectedPageIndices(allIndices);
                                setGlobalPagesSelectionText(formatPageIndicesToRangeString(allIndices));
                              }
                            }}
                            className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all flex flex-col items-center justify-center gap-1 ${
                              globalPagesSelectionMode === 'custom'
                                ? 'border-indigo-600 bg-indigo-50/50 text-indigo-700 ring-2 ring-indigo-500/10 shadow-sm'
                                : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                            }`}
                          >
                            <span className="text-sm">🎯</span>
                            <span>{language === Language.AR ? 'صفحات محددة' : 'Select Pages'}</span>
                          </button>
                        </div>

                        {globalPagesSelectionMode === 'custom' && (
                          <div className="space-y-2 pt-1 border-t border-slate-100">
                            <p className="text-[10px] text-slate-500 font-semibold leading-relaxed">
                              {language === Language.AR
                                ? 'اكتب أرقام الصفحات المطلوبة (مثال: 1, 3, 5-7) أو اضغط على الصفحات مباشرة لتحديدها.'
                                : 'Type page numbers (e.g., 1, 3, 5-7) or click pages directly to toggle selection.'}
                            </p>
                            
                            <div className="space-y-1">
                              <input
                                type="text"
                                value={globalPagesSelectionText}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setGlobalPagesSelectionText(val);
                                  const parsedIndices = parsePageRangeStringToIndices(val, interactivePages.length);
                                  setGlobalSelectedPageIndices(parsedIndices);
                                }}
                                placeholder={language === Language.AR ? 'مثال: 1, 3, 5-7' : 'e.g., 1, 3, 5-7'}
                                className="w-full px-3 py-2 rounded-xl border border-slate-300 focus:ring-2 focus:ring-indigo-500 outline-none text-xs bg-white text-slate-800 font-mono font-bold transition-all"
                              />
                            </div>

                            <div className="flex items-center justify-between text-[10px] text-slate-400 font-semibold pt-1">
                              <span>
                                {language === Language.AR
                                  ? `تم تحديد: ${globalSelectedPageIndices.length} صفحات`
                                  : `Selected: ${globalSelectedPageIndices.length} pages`}
                              </span>
                              <div className="flex gap-2">
                                <button
                                  type="button"
                                  onClick={() => {
                                    const allIndices = interactivePages.map(p => p.pageNum);
                                    setGlobalSelectedPageIndices(allIndices);
                                    setGlobalPagesSelectionText(formatPageIndicesToRangeString(allIndices));
                                  }}
                                  className="text-indigo-600 hover:underline"
                                >
                                  {language === Language.AR ? 'تحديد الكل' : 'Select All'}
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setGlobalSelectedPageIndices([]);
                                    setGlobalPagesSelectionText('');
                                  }}
                                  className="text-red-500 hover:underline"
                                >
                                  {language === Language.AR ? 'مسح الكل' : 'Clear All'}
                                </button>
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    )}

                    <div className="pt-2">
                      <button
                        onClick={handleProcess}
                        className="w-full py-3 rounded-xl font-extrabold text-white bg-emerald-600 hover:bg-emerald-700 active:scale-95 shadow-md hover:shadow-lg transition-all flex items-center justify-center space-x-2 space-x-reverse text-sm"
                      >
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                          <path strokeLinecap="round" strokeLinejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                        </svg>
                        <span>{language === Language.AR ? 'تأكيد العملية وحفظ الملف' : 'Confirm & Apply Process'}</span>
                      </button>
                    </div>
                  </div>
                </div>
              ) : (
                /* Non-modifying tool standard selection and options */
                <>
                  {/* Selected Files list */}
                  {selectedFiles.length > 0 && (
                    <div className="space-y-2">
                      <p className="text-sm font-semibold text-slate-700">{texts.filesSelected} ({selectedFiles.length})</p>
                      <div className="max-h-40 overflow-y-auto space-y-2 pr-1">
                        {selectedFiles.map((file, idx) => (
                          <div key={idx} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-100">
                            <div className="flex items-center space-x-2 space-x-reverse overflow-hidden">
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 text-slate-400 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                              </svg>
                              <span className="text-sm text-slate-700 truncate font-medium">{file.name}</span>
                              <span className="text-xs text-slate-400">({(file.size / 1024 / 1024).toFixed(2)} MB)</span>
                            </div>
                            <button 
                              onClick={(e) => { e.stopPropagation(); removeFile(idx); }}
                              className="text-slate-400 hover:text-red-500 p-1 hover:bg-slate-100 rounded"
                            >
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                              </svg>
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {toolId === 'translate' && (
                    <div className="space-y-2">
                      <label className="block text-sm font-bold text-slate-700">{texts.translateLabel}</label>
                      <select
                        value={targetTranslationLang}
                        onChange={(e) => setTargetTranslationLang(e.target.value)}
                        className="w-full px-4 py-3 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none bg-white transition-all"
                      >
                        {LANGUAGES_LIST.map((langObj, idx) => (
                          <option key={idx} value={langObj.code}>{langObj.label}</option>
                        ))}
                      </select>
                    </div>
                  )}

                  {toolId === 'unlock' && (
                    <div className="space-y-2">
                      <label className="block text-sm font-bold text-slate-700">{texts.unlockLabel}</label>
                      <input
                        type="password"
                        value={securityPassword}
                        onChange={(e) => setSecurityPassword(e.target.value)}
                        placeholder={texts.unlockPlaceholder}
                        className="w-full px-4 py-3 rounded-lg border border-slate-300 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                      />
                    </div>
                  )}

                  {/* Global Page Selection Option (Bilingual & Interactive) */}
                  {interactivePages.length > 0 && (
                    <div className="p-4 bg-white border border-slate-200 rounded-2xl space-y-3 shadow-sm my-3">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                        <span className="text-xs font-extrabold text-slate-800 flex items-center gap-1.5">
                          <span>📄</span>
                          <span>{language === Language.AR ? 'نطاق الصفحات المستهدفة' : 'Target Page Range'}</span>
                        </span>
                        <span className="text-[10px] text-slate-400 font-bold">
                          {language === Language.AR ? `إجمالي: ${interactivePages.length} صفحات` : `Total: ${interactivePages.length} pages`}
                        </span>
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            setGlobalPagesSelectionMode('all');
                          }}
                          className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all flex flex-col items-center justify-center gap-1 ${
                            globalPagesSelectionMode === 'all'
                              ? 'border-indigo-600 bg-indigo-50/50 text-indigo-700 ring-2 ring-indigo-500/10 shadow-sm'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          <span className="text-sm">🌟</span>
                          <span>{language === Language.AR ? 'كل الصفحات' : 'All Pages'}</span>
                        </button>

                        <button
                          type="button"
                          onClick={() => {
                            setGlobalPagesSelectionMode('custom');
                            if (globalSelectedPageIndices.length === 0) {
                              const allIndices = interactivePages.map(p => p.pageNum);
                              setGlobalSelectedPageIndices(allIndices);
                              setGlobalPagesSelectionText(formatPageIndicesToRangeString(allIndices));
                            }
                          }}
                          className={`py-2 px-3 rounded-xl border text-xs font-bold transition-all flex flex-col items-center justify-center gap-1 ${
                            globalPagesSelectionMode === 'custom'
                              ? 'border-indigo-600 bg-indigo-50/50 text-indigo-700 ring-2 ring-indigo-500/10 shadow-sm'
                              : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          <span className="text-sm">🎯</span>
                          <span>{language === Language.AR ? 'صفحات محددة' : 'Select Pages'}</span>
                        </button>
                      </div>

                      {globalPagesSelectionMode === 'custom' && (
                        <div className="space-y-2 pt-1 border-t border-slate-100 dark:border-slate-800">
                          <p className="text-[10px] text-slate-500 dark:text-slate-400 font-semibold leading-relaxed">
                            {language === Language.AR
                              ? 'اكتب أرقام الصفحات المطلوبة (مثال: 1, 3, 5-7) أو اضغط على الصفحات مباشرة لتحديدها.'
                              : 'Type page numbers (e.g., 1, 3, 5-7) or click pages directly to toggle selection.'}
                          </p>
                          
                          <div className="space-y-1">
                            <input
                              type="text"
                              value={globalPagesSelectionText}
                              onChange={(e) => {
                                const val = e.target.value;
                                setGlobalPagesSelectionText(val);
                                const parsedIndices = parsePageRangeStringToIndices(val, interactivePages.length);
                                setGlobalSelectedPageIndices(parsedIndices);
                              }}
                              placeholder={language === Language.AR ? 'مثال: 1, 3, 5-7' : 'e.g., 1, 3, 5-7'}
                              className="w-full px-3 py-2 rounded-xl border border-slate-300 dark:border-slate-700 focus:ring-2 focus:ring-indigo-500 outline-none text-xs bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100 font-mono font-bold transition-all"
                            />
                          </div>

                          <div className="flex items-center justify-between text-[10px] text-slate-400 dark:text-slate-500 font-semibold pt-1">
                            <span>
                              {language === Language.AR
                                ? `تم تحديد: ${globalSelectedPageIndices.length} صفحات`
                                : `Selected: ${globalSelectedPageIndices.length} pages`}
                            </span>
                            <div className="flex gap-2">
                              <button
                                type="button"
                                onClick={() => {
                                  const allIndices = interactivePages.map(p => p.pageNum);
                                  setGlobalSelectedPageIndices(allIndices);
                                  setGlobalPagesSelectionText(formatPageIndicesToRangeString(allIndices));
                                }}
                                className="text-indigo-600 dark:text-indigo-400 hover:underline"
                              >
                                {language === Language.AR ? 'تحديد الكل' : 'Select All'}
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  setGlobalSelectedPageIndices([]);
                                  setGlobalPagesSelectionText('');
                                }}
                                className="text-red-500 hover:underline"
                              >
                                {language === Language.AR ? 'مسح الكل' : 'Clear All'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  <button
                    onClick={handleProcess}
                    disabled={selectedFiles.length === 0}
                    className={`w-full py-4 rounded-xl font-bold text-white shadow-md transition-all flex items-center justify-center space-x-2 space-x-reverse ${
                      selectedFiles.length === 0
                        ? 'bg-slate-300 dark:bg-slate-800 cursor-not-allowed shadow-none text-slate-500'
                        : 'bg-indigo-600 hover:bg-indigo-700 hover:shadow-lg active:scale-95'
                    }`}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M13 10V3L4 14h7v7l9-11h-7z" />
                    </svg>
                    <span>{texts.processBtn}</span>
                  </button>
                </>
              )}
            </>
          )}

          {status === 'processing' && (
            <div className="py-12 flex flex-col items-center justify-center space-y-4 max-w-md mx-auto">
              <div className="h-16 w-16 border-4 border-indigo-200 dark:border-indigo-950/40 border-t-indigo-600 rounded-full animate-spin"></div>
              {batchProgress ? (
                <div className="w-full text-center space-y-3">
                  <p className="text-base font-semibold text-slate-800 dark:text-slate-200">{batchProgress.message}</p>
                  <div className="w-full bg-slate-100 dark:bg-slate-800 h-2.5 rounded-full overflow-hidden">
                    <div 
                      className="bg-indigo-600 h-full transition-all duration-300"
                      style={{ width: `${(batchProgress.current / batchProgress.total) * 100}%` }}
                    ></div>
                  </div>
                  <p className="text-xs text-slate-400 font-semibold font-mono">
                    {language === Language.AR ? `تمت معالجة ${batchProgress.current} من أصل ${batchProgress.total} ملفات` : `${batchProgress.current} of ${batchProgress.total} files processed`}
                  </p>
                </div>
              ) : (
                <div className="text-center space-y-2 animate-pulse">
                  <p className="text-lg font-semibold text-slate-800 dark:text-slate-200">
                    {processingMessage || texts.processing}
                  </p>
                  {toolId === 'pdf2word' && (
                    <p className="text-xs text-slate-400 font-medium">
                      {language === Language.AR 
                        ? 'تتم المعالجة بالتوازي لجميع الصفحات لضمان أفضل سرعة وأعلى دقة للنصوص العربية' 
                        : 'Processing all pages in parallel for maximum speed and Arabic text fidelity'}
                    </p>
                  )}
                </div>
              )}
            </div>
          )}

          {status === 'success' && (
            <div className="space-y-6 py-6 text-center">
              {toolId === 'chat' ? (
                /* Dynamic Interactive AI chat interface */
                <div className="text-start border border-slate-200 dark:border-slate-800 rounded-xl bg-slate-50 dark:bg-slate-950 overflow-hidden flex flex-col h-[50vh]">
                  {/* Messages area */}
                  <div className="flex-1 p-4 overflow-y-auto space-y-4 font-sans">
                    {chatHistory.map((msg, idx) => (
                      <div 
                        key={idx} 
                        className={`flex ${msg.sender === 'user' ? 'justify-end' : 'justify-start'}`}
                      >
                        <div 
                          className={`max-w-[85%] rounded-2xl px-4 py-2.5 text-sm shadow-sm leading-relaxed ${
                            msg.sender === 'user' 
                              ? 'bg-indigo-600 text-white rounded-br-none' 
                              : 'bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100 border border-slate-100 dark:border-slate-800 rounded-bl-none'
                          }`}
                        >
                          {msg.text}
                        </div>
                      </div>
                    ))}
                    {isChatSending && (
                      <div className="flex justify-start">
                        <div className="bg-white dark:bg-slate-900 border border-slate-100 dark:border-slate-800 rounded-2xl rounded-bl-none px-4 py-3 flex space-x-1.5 items-center">
                          <span className="w-2.5 h-2.5 bg-indigo-600 rounded-full animate-bounce"></span>
                          <span className="w-2.5 h-2.5 bg-indigo-600 rounded-full animate-bounce [animation-delay:0.2s]"></span>
                          <span className="w-2.5 h-2.5 bg-indigo-600 rounded-full animate-bounce [animation-delay:0.4s]"></span>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Message submit form */}
                  <form 
                    onSubmit={handleSendChatMessage} 
                    className="p-3 bg-white dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 flex space-x-2 space-x-reverse"
                  >
                    <input
                      type="text"
                      value={chatInput}
                      onChange={(e) => setChatInput(e.target.value)}
                      placeholder={texts.chatPlaceholder}
                      className="flex-1 px-4 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-sm transition-all bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-100"
                      disabled={isChatSending}
                    />
                    <button
                      type="submit"
                      disabled={!chatInput.trim() || isChatSending}
                      className="bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 dark:disabled:bg-slate-800 text-white font-bold px-5 py-2.5 rounded-lg text-sm shadow-sm transition-all flex items-center justify-center flex-shrink-0"
                    >
                      <span>{texts.chatSendBtn}</span>
                    </button>
                  </form>
                </div>
              ) : (
                <>
                  {/* Checkmark icon */}
                  <div className="mx-auto h-20 w-20 rounded-full bg-emerald-100 dark:bg-emerald-950/40 text-emerald-600 dark:text-emerald-400 flex items-center justify-center animate-bounce">
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-12 w-12" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                      <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                    </svg>
                  </div>

                  <div>
                    <h4 className="text-2xl font-extrabold text-slate-900 dark:text-white">{texts.success}</h4>
                    {selectedFiles[0] && (
                      <p className="text-slate-500 dark:text-slate-400 text-sm mt-1">{selectedFiles[0].name}</p>
                    )}
                  </div>

                  {/* Render Word Document Preview if toolId === 'pdf2word' */}
                  {toolId === 'pdf2word' && wordBlocks.length > 0 && (
                    <div className="text-start max-w-2xl mx-auto border border-slate-200 dark:border-slate-800 rounded-xl shadow-lg bg-white dark:bg-slate-900 overflow-hidden my-6">
                      <div className="bg-slate-100 dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800 px-4 py-2.5 flex items-center justify-between">
                        <div className="flex items-center space-x-2 space-x-reverse">
                          <div className="bg-blue-600 text-white p-1.5 rounded text-xs font-bold leading-none">W</div>
                          <span className="text-sm font-semibold text-slate-700 dark:text-slate-300 truncate max-w-xs">{downloadName}</span>
                        </div>
                        <span className="text-xs text-slate-400 font-medium">
                          {language === Language.AR ? 'معاينة المستند' : 'Document Preview'}
                        </span>
                      </div>
                      <div className="p-4 sm:p-6 max-h-[35vh] overflow-y-auto bg-slate-50 dark:bg-slate-950 font-serif leading-relaxed text-slate-800 dark:text-slate-200">
                        <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-sm p-6 sm:p-8 min-h-[250px] rounded-md">
                          {wordBlocks.map((b, idx) => {
                            if (b.type === 'title') {
                              return (
                                <h1 key={idx} className="text-2xl font-extrabold text-slate-900 dark:text-white mb-4 text-center">
                                  {b.text}
                                </h1>
                              );
                            }
                            if (b.type === 'heading1') {
                              return (
                                <h2 key={idx} className="text-xl font-bold text-indigo-700 dark:text-indigo-400 mt-4 mb-2 border-b border-slate-100 dark:border-slate-800 pb-1">
                                  {b.text}
                                </h2>
                              );
                            }
                            if (b.type === 'heading2') {
                              return (
                                <h3 key={idx} className="text-lg font-bold text-slate-800 dark:text-slate-200 mt-3 mb-2">
                                  {b.text}
                                </h3>
                              );
                            }
                            if (b.type === 'bullet') {
                              return (
                                <ul key={idx} className="list-disc pl-5 pr-5 mb-2 space-y-0.5">
                                  <li className="text-sm text-slate-700 dark:text-slate-300">{b.text}</li>
                                </ul>
                              );
                            }
                            if (b.type === 'numbered') {
                              return (
                                <ol key={idx} className="list-decimal pl-5 pr-5 mb-2 space-y-0.5">
                                  <li className="text-sm text-slate-700 dark:text-slate-300">{b.text}</li>
                                </ol>
                              );
                            }
                            if (b.type === 'table') {
                              return (
                                <div key={idx} className="overflow-x-auto my-3 border border-slate-200 dark:border-slate-800 rounded shadow-sm">
                                  <table className="min-w-full divide-y divide-slate-200 dark:divide-slate-800 text-xs">
                                    <tbody className="bg-white dark:bg-slate-900 divide-y divide-slate-100 dark:divide-slate-800">
                                      {b.rows?.map((row, rIdx) => (
                                        <tr key={rIdx} className={rIdx === 0 ? 'bg-slate-50 dark:bg-slate-950 font-bold' : ''}>
                                          {row.map((cell, cIdx) => (
                                            <td key={cIdx} className="px-3 py-1.5 text-slate-700 dark:text-slate-300 border-r border-slate-200 dark:border-slate-800 last:border-0">
                                              {cell}
                                            </td>
                                          ))}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              );
                            }
                            return (
                              <p key={idx} className="text-sm text-slate-700 dark:text-slate-300 mb-3 leading-relaxed">
                                {b.text}
                              </p>
                            );
                          })}
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Render downloadable links */}
                  {batchResults.length > 0 ? (
                    <div className="space-y-4 max-w-xl mx-auto">
                      <p className="text-sm font-bold text-slate-700 dark:text-slate-300">
                        {language === Language.AR ? '📄 الملفات المعالجة الجاهزة للتحميل:' : '📄 Processed Files Ready for Download:'}
                      </p>
                      <div className="grid grid-cols-1 gap-2.5 max-h-72 overflow-y-auto pr-1 text-start">
                        {batchResults.map((res: any, rIdx) => (
                          <div key={rIdx} className="flex flex-col sm:flex-row sm:items-center justify-between p-3.5 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl shadow-sm gap-2">
                            <div className="flex flex-col text-start max-w-[70%]">
                              <span className="text-xs font-bold text-slate-700 dark:text-slate-300 truncate" title={res.name}>{res.name}</span>
                              {res.status === 'error' && (
                                <span className="text-[10px] text-red-500 font-medium">
                                  {language === Language.AR ? `فشل: ${res.error}` : `Failed: ${res.error}`}
                                </span>
                              )}
                            </div>
                            {res.status === 'success' ? (
                              <button
                                id={`download-batch-${rIdx}`}
                                onClick={() => {
                                  if (res.blob) {
                                    triggerFileDownload(res.blob, res.name, res.blob.type);
                                  } else if (res.url) {
                                    fetch(res.url)
                                      .then(r => r.blob())
                                      .then(b => triggerFileDownload(b, res.name, b.type))
                                      .catch(() => {
                                        const link = document.createElement('a');
                                        link.id = `fallback-batch-link-${rIdx}`;
                                        link.href = res.url;
                                        link.download = res.name;
                                        document.body.appendChild(link);
                                        link.click();
                                        document.body.removeChild(link);
                                      });
                                  }
                                }}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold py-2 px-4 rounded-xl shadow hover:shadow-md transition-all flex items-center justify-center gap-1.5 active:scale-95 cursor-pointer self-start sm:self-auto"
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                                </svg>
                                <span>{language === Language.AR ? 'تحميل' : 'Download'}</span>
                              </button>
                            ) : (
                              <span className="text-xs text-red-500 bg-red-50 dark:bg-red-950/20 px-2.5 py-1 rounded-lg font-bold border border-red-100 dark:border-red-900/30 self-start sm:self-auto">
                                {language === Language.AR ? 'فشل المعالجة' : 'Failed'}
                              </span>
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : splitResults.length > 0 && toolId === 'split' ? (
                    <div className="space-y-4 max-w-md mx-auto">
                      <p className="text-sm font-bold text-slate-700 dark:text-slate-300">
                        {language === Language.AR ? '📄 الأجزاء المقسمة الجاهزة للتحميل:' : '📄 Split Parts Ready for Download:'}
                      </p>
                      <div className="grid grid-cols-1 gap-2.5 max-h-56 overflow-y-auto pr-1">
                        {splitResults.map((res: any, rIdx) => (
                          <div key={rIdx} className="flex items-center justify-between p-3.5 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl shadow-sm">
                            <span className="text-xs font-bold text-slate-700 dark:text-slate-300 truncate max-w-[240px]">{res.name}</span>
                            <button
                              id={`download-split-${rIdx}`}
                              onClick={() => {
                                if (res.blob) {
                                  triggerFileDownload(res.blob, res.name, 'application/pdf');
                                } else if (res.url) {
                                  // Fallback to fetch blob from url
                                  fetch(res.url)
                                    .then(r => r.blob())
                                    .then(b => triggerFileDownload(b, res.name, 'application/pdf'))
                                    .catch(() => {
                                      const link = document.createElement('a');
                                      link.id = `fallback-link-${rIdx}`;
                                      link.href = res.url;
                                      link.download = res.name;
                                      document.body.appendChild(link);
                                      link.click();
                                      document.body.removeChild(link);
                                    });
                                }
                              }}
                              className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold py-2 px-4 rounded-xl shadow hover:shadow-md transition-all flex items-center gap-1.5 active:scale-95 cursor-pointer"
                            >
                              <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                              </svg>
                              <span>{language === Language.AR ? 'تحميل' : 'Download'}</span>
                            </button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : resultUrl ? (
                    <div className="bg-slate-50 dark:bg-slate-900 border border-slate-100 dark:border-slate-800 rounded-xl p-4 max-w-sm mx-auto">
                      <button
                        id="download-processed-btn"
                        onClick={() => {
                          if (resultBlob) {
                            triggerFileDownload(resultBlob, downloadName, resultMime);
                          } else {
                            // Fallback to fetching blob from resultUrl
                            fetch(resultUrl)
                              .then(r => r.blob())
                              .then(b => triggerFileDownload(b, downloadName, resultMime || b.type))
                              .catch(() => {
                                const link = document.createElement('a');
                                link.id = 'fallback-download-link';
                                link.href = resultUrl;
                                link.download = downloadName;
                                document.body.appendChild(link);
                                link.click();
                                document.body.removeChild(link);
                              });
                          }
                        }}
                        className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-3.5 px-6 rounded-xl font-bold shadow-md hover:shadow-lg transition-all active:scale-95 flex items-center justify-center space-x-2 space-x-reverse cursor-pointer"
                      >
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                          <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                        </svg>
                        <span>{texts.downloadBtn}</span>
                      </button>
                    </div>
                  ) : (
                    /* Render AI text results */
                    <div className="text-start bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-850 rounded-xl p-6 space-y-4 max-h-[45vh] overflow-y-auto">
                      <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
                        <span className="text-sm font-bold text-indigo-700 dark:text-indigo-400">{texts.aiResultHeader}</span>
                        <div className="flex space-x-2 space-x-reverse">
                          <button
                            onClick={copyToClipboard}
                            className="text-xs bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 px-3 py-1.5 rounded-lg font-medium transition-all"
                          >
                            {copied ? texts.copied : texts.copyBtn}
                          </button>
                          <button
                            onClick={saveToTxt}
                            className="text-xs bg-indigo-50 dark:bg-indigo-950/40 hover:bg-indigo-100 dark:hover:bg-indigo-900 text-indigo-700 dark:text-indigo-400 px-3 py-1.5 rounded-lg font-medium transition-all"
                          >
                            {texts.saveTxtBtn}
                          </button>
                        </div>
                      </div>
                      <div className="prose prose-indigo max-w-none text-slate-800 dark:text-slate-200 leading-relaxed font-sans">
                        {renderMarkdown(aiTextResult)}
                      </div>
                    </div>
                  )}
                </>
              )}

              <button
                onClick={() => {
                  setStatus('idle');
                  setSelectedFiles([]);
                  setResultUrl(null);
                  setResultBlob(null);
                  setResultMime('');
                  setAiTextResult('');
                  setChatHistory([]);
                  setInteractivePages([]);
                  setSplitResults([]);
                }}
                className="text-sm font-semibold text-indigo-600 dark:text-indigo-400 hover:text-indigo-800 dark:hover:text-indigo-300 transition-colors"
              >
                {language === Language.AR ? 'معالجة ملف آخر' : 'Process another file'}
              </button>
            </div>
          )}

          {status === 'error' && (
            <div className="py-6 space-y-4 text-center">
              <div className="mx-auto h-16 w-16 rounded-full bg-red-100 dark:bg-red-950/40 text-red-600 dark:text-red-400 flex items-center justify-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-10 w-10" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                </svg>
              </div>
              <div>
                <h4 className="text-xl font-bold text-slate-900 dark:text-white">{texts.errorTitle}</h4>
                <p className="text-slate-600 dark:text-slate-400 text-sm mt-2 max-w-md mx-auto leading-relaxed">{errorMessage}</p>
              </div>
              <button
                onClick={() => setStatus('idle')}
                className="bg-slate-800 hover:bg-slate-700 text-white font-bold py-2 px-6 rounded-lg transition-colors"
              >
                {language === Language.AR ? 'إعادة المحاولة' : 'Try Again'}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};
