export const highlightCurrentNav = (linkName:any) => {
    let navLinks = document.querySelectorAll('.nav-click');
    if (navLinks.length > 0) {
        navLinks.forEach(item => {
            item.classList.remove("nav-click");
        });
    }
    if(document.getElementById(linkName)!=null){
        document.getElementById(linkName)?.classList.add('nav-click');
    }
};

export const sortDataByTitle = (data: any[],property: any)=>{
    data.sort((a,b)=>{
        let fa = a[`${property}`].toLowerCase();
         let fb = b[`${property}`].toLowerCase();
       if (fa < fb) {
               return -1;
           }
           if (fa > fb) {
               return 1;
           }
           return 0;
       });
       return data;
};
