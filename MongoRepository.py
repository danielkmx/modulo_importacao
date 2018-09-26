from flask import Flask, jsonify
from flask_cors import CORS
import KoboApiService as kapi
from flask_restplus import Resource, Api

app = Flask(__name__)
api = Api(app)
CORS(app)
app.config['JSON_AS_ASCII'] = False

ns = api.namespace('Formularios', description='Operações para filtrar dados das enquetes registradas no KoboToolbox')

@ns.route('/<string:url>/<int:i_d>/<string:username>/<string:password>')
class Formulario(Resource):
    def get(self, url, i_d, username, password):
        """ Retorna todas enquetes preenchidas

            Use a rota /formularios para conseguir o ID das enquetes



                 Urls do kobo legado:
                kc.humanitarianresponse.info
                kobocatdocker.kobo.techo.org

                """

        return jsonify(kapi.retorna_respostas_com_labels(url, i_d, username, password))
@ns.route('/<string:url>/<string:username>/<string:password>')
class Formularios(Resource):
    def get(self, url, username, password):
        """
        Retorna Ids e nomes de todos formularios

             Urls do kobo legado:


                kc.humanitarianresponse.info
                kobocat.docker.kobo.techo.org
            """
        return jsonify(kapi.imprimir_lista_formularios(url, username, password))

@ns.route('/perguntas/<string:url>/<int:i_d>/<string:username>/<string:password>')
class Formularios(Resource):
    def get(self, url, i_d, username, password):
        """
        Retorna Ids e nomes de todos formularios

             Urls do kobo legado:


                kc.humanitarianresponse.info
                kobocat.docker.kobo.techo.org
            """
        return jsonify(kapi.retorna_lista_perguntas(url, i_d, username, password))


if __name__ == '__main__':
    app.run(debug=True)

