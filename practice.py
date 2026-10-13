import numpy
print('numpy version', numpy.__version__)
import matplotlib 
import sklearn
import pandas
import scipy 

from pandas.plotting import scatter_matrix
from matplotlib import pyplot as plt
from sklearn.datasets import load_iris
from sklearn.model_selection import train_test_split
from sklearn.model_selection import StratifiedKFold
from sklearn.model_selection import cross_val_score
from sklearn.metrics import classification_report
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.naive_bayes import GaussianNB
from sklearn.svm import SVC

names = ['sepal-length', 'sepal-width', 'petal-length','petal-width','class']
iris = load_iris()
dataset = pandas.DataFrame(iris.data, columns=names[:4])
dataset['class'] = pandas.Series(iris.target).map(
    lambda target: f'Iris-{iris.target_names[target]}'
)

print(dataset.shape)
print(dataset.head())
print(dataset.describe())
print(dataset.groupby('class').size())





scatter_matrix(dataset)
plt.show()


array = dataset.values
X = array[:,0:4]
Y = array[:,4]
X_train, X_validation, Y_train, Y_validation = train_test_split(X, Y, test_size=0.20, random_state=1)

models = []
models.append(('LR', LogisticRegression(max_iter=200)))
models.append(('LDA',LinearDiscriminantAnalysis()))
models.append(('KNN', KNeighborsClassifier()))
models.append(('SVM',SVC(gamma='auto')))
models.append(('CART',DecisionTreeClassifier()))
models.append(('NB', GaussianNB()))

results = []
names = []

for name,model in models:
    kfold = StratifiedKFold(n_splits=10, random_state=1, shuffle=True)
    cv_results = cross_val_score(model, X_train, Y_train, cv=kfold, scoring='accuracy')
    results.append(cv_results)
    names.append(name)
    print('%s, %f(%f)' % (name, cv_results.mean(), cv_results.std()))


model = SVC(gamma='auto')
model.fit(X_train, Y_train)
predictions = model.predict(X_validation)

print(predictions)
print(accuracy_score(Y_validation, predictions))
print(confusion_matrix(Y_validation, predictions))