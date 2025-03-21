from graphviz import Digraph

# Criando o diagrama UML
uml = Digraph('RabbitMQ_Email_System', filename='rabbitmq_email_system', format='png')

# Configuração geral
uml.attr(rankdir='TB', size='10')

# Classes principais
uml.node('EmailSender', '<<interface>>\nEmailSender', shape='interface', style='filled', fillcolor='lightgrey')
uml.node('EmailProducer', 'EmailProducer', shape='rectangle', style='filled', fillcolor='lightblue')
uml.node('EmailConsumer', 'EmailConsumer', shape='rectangle', style='filled', fillcolor='lightblue')
uml.node('EmailDLQProcessor', 'EmailDLQProcessor', shape='rectangle', style='filled', fillcolor='lightblue')
uml.node('RabbitMQConfig', 'RabbitMQConfig', shape='rectangle', style='filled', fillcolor='lightgrey')
uml.node('EmailMessage', 'EmailMessage', shape='rectangle', style='filled', fillcolor='lightgreen')
uml.node('OrderService', 'OrderService', shape='rectangle', style='filled', fillcolor='lightyellow')
uml.node('OrderController', 'OrderController', shape='rectangle', style='filled', fillcolor='lightyellow')

# Relacionamentos
uml.edge('EmailSender', 'EmailConsumer', arrowhead='onormal', label='implements')
uml.edge('EmailProducer', 'RabbitMQConfig', arrowhead='vee', label='uses')
uml.edge('EmailConsumer', 'RabbitMQConfig', arrowhead='vee', label='listens to')
uml.edge('EmailConsumer', 'EmailMessage', arrowhead='vee', label='processes')
uml.edge('EmailDLQProcessor', 'RabbitMQConfig', arrowhead='vee', label='listens to DLQ')
uml.edge('OrderService', 'EmailProducer', arrowhead='vee', label='calls')
uml.edge('OrderController', 'OrderService', arrowhead='vee', label='calls')

# Gerando a imagem
uml_path = '/mnt/data/rabbitmq_email_system.png'
uml.render(uml_path, format='png', cleanup=True)

# Retornando o caminho da imagem gerada
uml_path
