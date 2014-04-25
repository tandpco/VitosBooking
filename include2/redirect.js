
function redirect(url, msg)
{
   var TARG_ID = "redirect";
   var DEF_MSG = "Redirecting...";

   if( ! msg )
   {
      msg = DEF_MSG;
   }

   if( ! url )
   {
      throw new Error('You didn\'t include the "url" parameter');
   }


   var e = document.getElementById(TARG_ID);

   if( ! e )
   {
      throw new Error('"redirect" element id not found');
   }

   var cTicks = parseInt(e.innerHTML);

//   var timer = setInterval(function()
//   {
//      // 2011-07-20 TAM: Grab value from div so countdown can be reset elsewhere
//      cTicks = parseInt(e.innerHTML);
//      
//      if( cTicks )
//      {
//         e.innerHTML = --cTicks;
//      }
//      else
//      {
//         clearInterval(timer);
//         document.body.innerHTML = msg;
//         location = url;	  
//      }
//
//   }, 1000);
}
