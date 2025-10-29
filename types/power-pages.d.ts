// Power Pages WebAPI TypeScript declarations
declare namespace webapi {
  interface AjaxOptions {
    type: string;
    url: string;
    contentType?: string;
    data?: string;
    success?: (res: any, status: string, xhr: XMLHttpRequest) => void;
    error?: (xhr: XMLHttpRequest, status: string, error: string) => void;
  }

  function safeAjax(options: AjaxOptions): void;
}

// Power Pages global objects
declare var $: any; // jQuery if available
