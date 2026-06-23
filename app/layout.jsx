import { Outfit } from "next/font/google";
import { Toaster } from "react-hot-toast";
import StoreProvider from "@/app/StoreProvider";
import "./globals.css";

const outfit = Outfit({ subsets: ["latin"], weight: ["400", "500", "600"] });

export const metadata = {
    title: "GoCart. - Shop smarter",
    description: "GoCart. - Shop smarter",
};

export default function RootLayout({ children }) {
    return (
        <html lang="en">
            <head>
                {/* HAM Voice Widget */}
                <script src="https://voicesaasbd2-production.up.railway.app/cdn/apex-voice.js"></script>
                <script dangerouslySetInnerHTML={{
                    __html: `
                        window.addEventListener('load', () => {
                            ApexVoice.init({
                                apiKey: "ham_5a18c296",
                                companyName: "gocart",
                                serverUrl: (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1' || window.location.hostname.startsWith('192.168.')) ? (window.location.hostname + ':8001') : 'voicesaasbd2-production.up.railway.app',
                                systemInstruction: "Tu es l'assistant IA GoCart. Aide l'utilisateur avec les produits du catalogue.\\nVoici les produits disponibles dans le catalogue avec leurs identifiants (IDs) :\\n- ID 'prod_1': Modern table lamp ($299)\\n- ID 'prod_2': Smart speaker gray ($349)\\n- ID 'prod_3': Smart watch white ($499)\\n- ID 'prod_4': Wireless headphones ($599)\\n- ID 'prod_5': Smart watch black ($399)\\n- ID 'prod_6': Security Camera ($450)\\n- ID 'prod_7': Smart Pen for iPad ($750)\\n- ID 'prod_8': Home Theater ($899)\\n- ID 'prod_9': Apple Wireless Earbuds ($799)\\n- ID 'prod_10': Apple Smart Watch ($1499)\\n- ID 'prod_11': RGB Gaming Mouse ($249)\\n- ID 'prod_12': Smart Home Cleaner ($1699)\\n\\nQuand l'utilisateur parle d'un produit, s'il souhaite le voir, avoir plus d'informations dessus ou aller sur sa page, tu DOIS appeler la fonction 'view_product' avec le 'productId' correspondant.",
                                tools: [
                                    {
                                        functionDeclarations: [
                                            {
                                                name: "view_product",
                                                description: "Redirige l'utilisateur vers la page de détails d'un produit spécifique pour l'afficher à l'écran.",
                                                parameters: {
                                                    type: "OBJECT",
                                                    properties: {
                                                        productId: {
                                                            type: "STRING",
                                                            description: "L'identifiant (ID) du produit à afficher, par exemple 'prod_1'."
                                                        }
                                                    },
                                                    required: ["productId"]
                                                }
                                            }
                                        ]
                                    }
                                ],
                                actions: {
                                     view_product: (args) => {
                                         if (args && args.productId) {
                                             const url = '/product/' + args.productId;
                                             if (window.nextRouter) {
                                                 window.nextRouter.push(url);
                                             } else {
                                                 window.location.href = url;
                                             }
                                         }
                                     },
                                     navigate_to_product: (args) => {
                                         if (args && args.product) {
                                             const target = args.product.toLowerCase().trim();
                                             if (target.startsWith('prod_')) {
                                                 const url = '/product/' + args.product;
                                                 if (window.nextRouter) {
                                                     window.nextRouter.push(url);
                                                 } else {
                                                     window.location.href = url;
                                                 }
                                                 return;
                                             }
                                             const productMap = {
                                                 "modern table lamp": "prod_1",
                                                 "smart speaker gray": "prod_2",
                                                 "smart watch white": "prod_3",
                                                 "wireless headphones": "prod_4",
                                                 "smart watch black": "prod_5",
                                                 "security camera": "prod_6",
                                                 "smart pen for ipad": "prod_7",
                                                 "home theater": "prod_8",
                                                 "apple wireless earbuds": "prod_9",
                                                 "apple smart watch": "prod_10",
                                                 "rgb gaming mouse": "prod_11",
                                                 "smart home cleaner": "prod_12"
                                             };
                                             let foundId = null;
                                             for (const [name, id] of Object.entries(productMap)) {
                                                 if (target.includes(name) || name.includes(target)) {
                                                     foundId = id;
                                                     break;
                                                 }
                                             }
                                             if (foundId) {
                                                 const url = '/product/' + foundId;
                                                 if (window.nextRouter) {
                                                     window.nextRouter.push(url);
                                                 } else {
                                                     window.location.href = url;
                                                 }
                                             }
                                         }
                                     }
                                }
                            });
                        });
                    `
                }} />
            </head>
            <body className={`${outfit.className} antialiased`}>
                <StoreProvider>
                    <Toaster />
                    {children}
                </StoreProvider>
            </body>
        </html>
    );
}
