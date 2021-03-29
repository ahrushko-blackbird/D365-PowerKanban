export const loadExternalScript = async (url: string) => {
    return new Promise((resolve, reject) => {
        const scriptTag = document.createElement("script");

        scriptTag.src = url;
        scriptTag.async = true;
        scriptTag.onload = resolve;

        document.body.appendChild(scriptTag);
    });
};

export const loadExternalResource = async (url: string) => {
    const content = await fetch(url);
    return content.text();
};